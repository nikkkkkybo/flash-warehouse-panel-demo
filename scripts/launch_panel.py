# -*- coding: utf-8 -*-
"""
智能启动器（根治「旧进程霸占端口 → 新代码永不启动」问题）。

启动流程（按端口 5000 起，最多顺延到 5002）：
  1) 端口空闲 → 直接启动面板（os.execv 原位替换，关闭窗口即停止服务）
  2) 端口被占 → 请求 http://127.0.0.1:<port>/api/version 比对磁盘代码指纹：
     a. 指纹一致   → 最新版已在运行，打开浏览器后退出（无需重启）
     b. 指纹不一致 → 旧版面板在运行 → 校验占用进程命令行确属 web_panel.py
                     → taskkill 结束旧进程 → 等端口释放 → 启动新代码
     c. 无版本接口 → 占用者不是本面板 → 顺延下一端口重试
  3) 全部端口被非面板程序占用 → 提示后放弃

指纹 = md5(web_panel.py) 前 12 位，代码一变指纹即变，精确区分新旧。
"""
import io
import os
import sys
import time
import socket
import hashlib
import subprocess
import urllib.request

HERE = os.path.dirname(os.path.abspath(__file__))
PANEL = os.path.join(HERE, "web_panel.py")
PORTS = [5000, 5001, 5002]


def disk_version():
    """磁盘代码指纹：与面板 PANEL_VERSION 同一函数算，杜绝两边口径漂移。"""
    from version_stamp import panel_version
    return panel_version()


def port_free(port):
    """端口是否空闲。"""
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        s.settimeout(1)
        return s.connect_ex(("127.0.0.1", port)) != 0


def remote_info(port, timeout=3):
    """运行中面板的 {version, pid}；连不上/超时 → None。

    pid 用来区分「正在服务的进程」与「同端口上残留的其他面板进程」：Windows 允许多个
    进程同时 bind 同一端口，残留进程不接流量却占着端口，一旦服务进程退出就可能接手
    旧代码——这正是「改了没生效」假象的根因。
    """
    try:
        with urllib.request.urlopen(
                f"http://127.0.0.1:{port}/api/version", timeout=timeout) as r:
            import json
            d = json.loads(r.read().decode("utf-8"))
            return {"version": d.get("version"), "pid": d.get("pid")}
    except Exception:
        return None


def remote_version(port, timeout=3):
    """运行中面板的版本指纹；无 /api/version、连接失败、超时 → None。"""
    info = remote_info(port, timeout=timeout)
    return info.get("version") if info else None


def port_owner_pids(port):
    """netstat 解析占用端口的全部 LISTENING PID（可能多个：IPv4/IPv6/双绑定）；无 → []。"""
    pids = []
    try:
        out = subprocess.run(["netstat", "-ano"], capture_output=True,
                             creationflags=subprocess.CREATE_NO_WINDOW).stdout.decode("gbk", "replace")
        for ln in out.splitlines():
            parts = ln.split()
            if len(parts) >= 5 and "LISTENING" in ln.upper():
                laddr = parts[1]
                if laddr.endswith(f":{port}"):
                    pid = parts[-1]
                    if pid.isdigit() and int(pid) not in pids:
                        pids.append(int(pid))
    except Exception:
        pass
    return pids


def port_owner_pid(port):
    """兼容旧名：第一个占用 PID 或 None。"""
    pids = port_owner_pids(port)
    return pids[0] if pids else None


def pid_cmdline(pid):
    """进程命令行；失败 → ''。"""
    try:
        out = subprocess.run(
            ["powershell", "-NoProfile", "-Command",
             f"(Get-CimInstance Win32_Process -Filter 'ProcessId={pid}').CommandLine"],
            capture_output=True, creationflags=subprocess.CREATE_NO_WINDOW
        ).stdout.decode("utf-8", "replace")
        return out.strip()
    except Exception:
        return ""


def wait_port_free(port, sec=6):
    for _ in range(int(sec / 0.25)):
        if port_free(port):
            return True
        time.sleep(0.25)
    return port_free(port)


def py_procs():
    """全部 python 进程 [(pid, ppid, cmdline)]；失败 → []。"""
    try:
        out = subprocess.run(
            ["powershell", "-NoProfile", "-Command",
             "Get-CimInstance Win32_Process -Filter \"Name='python.exe'\" | "
             "ForEach-Object { \"$($_.ProcessId)`t$($_.ParentProcessId)`t$($_.CommandLine)\" }"],
            capture_output=True, creationflags=subprocess.CREATE_NO_WINDOW).stdout
        rows = []
        for ln in out.decode("gbk", "replace").splitlines():
            parts = ln.split("\t", 2)
            if len(parts) >= 2 and parts[0].strip().isdigit():
                rows.append((int(parts[0].strip()),
                             int(parts[1].strip()) if parts[1].strip().isdigit() else 0,
                             (parts[2] if len(parts) > 2 else "").strip()))
        return rows
    except Exception:
        return []


def retire_panel_procs(keep=()):
    """清理重复的面板实例。

    为什么必须连守护一起杀：run_guarded 会在子面板被结束后 5 秒内重新拉起，
    只杀 web_panel 等于「杀一个长一个」——实测 5000 端口上因此积到 3 个监听进程，
    每个都完整持有自家总表，员工机（4G 起）内存直接翻倍。顺序：先守护后子进程。
    keep 中的 pid 及其守护祖先保留。
    """
    procs = py_procs()
    if not procs:
        return []
    by_pid = {p[0]: p for p in procs}
    keep = set(int(k) for k in keep if k)
    # 保留链：keep 自身 + 其祖先（守护它的 launch_panel）
    keep_tree = set(keep)
    for k in list(keep_tree):
        cur = k
        while cur in by_pid and by_pid[cur][1] and by_pid[cur][1] not in keep_tree:
            cur = by_pid[cur][1]
            if cur in keep_tree or cur not in by_pid:
                break
            keep_tree.add(cur)

    def _role(cmd):
        if "web_panel.py" in cmd:
            return "panel"
        if "launch_panel.py" in cmd:
            return "guard"
        return ""

    victims = [(pid, _role(cmd)) for pid, ppid, cmd in procs
               if pid not in keep_tree and _role(cmd)]
    killed = []
    for role in ("guard", "panel"):          # 先断守护，再杀面板
        for pid, r in victims:
            if r != role or pid in killed:
                continue
            subprocess.run(["taskkill", "/PID", str(pid), "/F"],
                           capture_output=True, creationflags=subprocess.CREATE_NO_WINDOW)
            killed.append(pid)
    return killed


def run_guarded(port):
    """守护式启动：面板崩溃自动拉起（连续 5 次异常退出后放弃，防崩溃循环）。
    正常退出（exit=0，如 Ctrl+C）不重启；关闭本窗口 = 整个控制台树终止，语义不变。"""
    os.environ["PANEL_PORT"] = str(port)
    fails = 0
    while True:
        p = subprocess.Popen([sys.executable, PANEL])
        try:
            code = p.wait()
        except KeyboardInterrupt:
            p.terminate()
            return 0
        if code == 0:
            print("[守护] 面板正常退出，不重启。")
            return 0
        fails += 1
        if fails >= 5:
            print(f"[守护] 面板连续 {fails} 次异常退出，放弃自动拉起（exit={code}）。")
            return code
        print(f"[守护] 面板异常退出（exit={code}），5 秒后第 {fails} 次拉起...")
        time.sleep(5)


def main():
    dv = disk_version()
    print(f"面板代码指纹: {dv}")
    for port in PORTS:
        if port_free(port):
            print(f"端口 {port} 空闲，守护式启动（崩溃自动拉起）...")
            return run_guarded(port)

        print(f"端口 {port} 已被占用，校验占用者...")
        rv = remote_version(port)
        pids = port_owner_pids(port)
        cmds = {p: pid_cmdline(p) for p in pids}
        # 任一占用进程命令行含 web_panel.py 即视为面板（旧版：指纹不一致或无版本接口）
        panel_pids = [p for p in pids if "web_panel.py" in cmds[p]]
        if rv == dv:
            # 指纹一致 = 最新代码在服务；但同端口可能残留不接流量的旧面板进程
            # （Windows 允许多进程 bind 同端口，且守护会自动拉起被杀的子进程）。
            # 保留「服务者 + 其守护 + 本进程」，其余重复实例全部清掉。
            info = remote_info(port) or {}
            serving = info.get("pid")
            killed = retire_panel_procs(keep={p for p in (serving, os.getpid()) if p})
            if killed:
                print(f"  [清理] 残留重复实例 {killed}（服务者 {serving}）")
            print(f"  最新版面板已在运行（指纹一致），直接打开浏览器")
            import webbrowser
            webbrowser.open(f"http://127.0.0.1:{port}")
            time.sleep(1)
            return
        if not panel_pids:
            print(f"  占用进程 {pids} 非本面板，顺延下一端口")
            continue
        # 旧版面板 → 连同其守护一起替换（只杀子进程会被守护 5 秒内拉起，越替换越多）
        print(f"  [替换] 旧版面板（PID {panel_pids}，指纹 {rv or '无版本接口'}）→ 结束并启动新版...")
        killed = retire_panel_procs(keep={os.getpid()})
        if killed:
            print(f"  [替换] 已结束 {killed}")
        if not wait_port_free(port):
            print(f"  端口 {port} 未能释放，顺延下一端口")
            continue
        return run_guarded(port)
    print("所有候选端口（5000-5002）均被非面板程序占用，无法启动。")
    print("请关闭占用程序后重试，或手工执行: python scripts/web_panel.py")
    try:
        if sys.stdin is not None and sys.stdin.isatty():
            input("按回车键退出...")   # 无控制台时不等人（否则会永久挂住）
    except Exception:
        pass


class _SilentStream:
    """无控制台（pythonw / 计划任务）时的 stdout/stderr 兜底。

    为什么必须兜：这些环境下 sys.stdout 是 None，直接 `sys.stdout.buffer` 会
    AttributeError 让启动器静默退出——实测就是这样把面板弄没的。print 丢进这里
    只是不显示，真正的运行痕迹在 output/logs/panel.log。
    """

    def write(self, *a):
        return 0

    def flush(self):
        pass

    def isatty(self):
        return False

    @property
    def buffer(self):
        return None


def _wrap_streams_utf8():
    for name in ("stdout", "stderr"):
        s = getattr(sys, name, None)
        if s is None or not hasattr(s, "buffer"):
            setattr(sys, name, _SilentStream())
            continue
        try:
            setattr(sys, name, io.TextIOWrapper(s.buffer, encoding="utf-8", errors="replace"))
        except Exception:
            pass


if __name__ == "__main__":
    _wrap_streams_utf8()
    main()
