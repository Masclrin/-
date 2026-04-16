#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
╔═══════════════════════════════════════════════════════════════════╗
║           原神挑战计时优化器  v1.0                                ║
║           Genshin Challenge Timer Optimizer                       ║
║                                                                   ║
║  基于Unity底层分析，利用以下原理优化挑战时间:                        ║
║    - 帧率锁定: 消除高刷显示器导致的帧级时间窃取                      ║
║    - 后台切换: 利用Unity失焦降频延迟客户端初始化                     ║
║    - CPU压力: 缓慢异步场景加载，拉平硬件差异                        ║
║    - 黑屏检测: 精确定位操作时机                                    ║
║                                                                   ║
║  用法:                                                            ║
║    pip install keyboard mss                                       ║
║    以管理员身份运行: python genshin_timer.py                      ║
║    在游戏中点击"开始挑战" → 立即按 F9 触发                        ║
║                                                                   ║
║  前置: RTSS 已安装并配置好帧率限制热键                             ║
╚═══════════════════════════════════════════════════════════════════╝
"""

import sys
import os
import time
import ctypes
import ctypes.wintypes
import threading
import signal
from dataclasses import dataclass
from multiprocessing import Process, Event as MpEvent
from typing import Optional, List

# ═══════════════════════════════════════════════════════════════
# 0. 环境与工具函数
# ═══════════════════════════════════════════════════════════════

_shutdown = threading.Event()          # 全局取消标志

user32 = ctypes.windll.user32
kernel32 = ctypes.windll.kernel32
SW_MINIMIZE, SW_RESTORE = 6, 9


def is_admin():
    """检查是否以管理员权限运行 (keyboard全局钩子需要)"""
    try:
        return bool(ctypes.windll.shell32.IsUserAnAdmin())
    except Exception:
        return False


def get_fg_hwnd():
    """获取当前前台窗口句柄"""
    return user32.GetForegroundWindow()


def win_minimize(hwnd):
    """最小化窗口 → Unity 收到 OnApplicationPause(true) + OnApplicationFocus(false)"""
    user32.ShowWindow(hwnd, SW_MINIMIZE)


def win_restore(hwnd):
    """恢复窗口并强制设为前台 (绕过 Windows SetForegroundWindow 限制)"""
    user32.ShowWindow(hwnd, SW_RESTORE)
    _force_fg(hwnd)


def _force_fg(hwnd):
    """
    Windows 安全机制限制: 只有前台进程才能 SetForegroundWindow.
    绕过方法: 用 AttachThreadInput 将当前线程附加到前台窗口线程.
    """
    if user32.SetForegroundWindow(hwnd):
        return True
    try:
        fhwnd = user32.GetForegroundWindow()
        ftid = user32.GetWindowThreadProcessId(fhwnd, None)
        ctid = kernel32.GetCurrentThreadId()
        if ftid != ctid:
            user32.AttachThreadInput(ftid, ctid, True)
            user32.ShowWindow(hwnd, SW_RESTORE)
            time.sleep(0.05)
            r = user32.SetForegroundWindow(hwnd)
            user32.AttachThreadInput(ftid, ctid, False)
            return bool(r)
    except Exception:
        pass
    return False


def beep(freq=1000, ms=100):
    """异步蜂鸣 (不阻塞主流程, 用作阶段提示音)"""
    def _b():
        try:
            import winsound
            winsound.Beep(freq, ms)
        except Exception:
            pass
    threading.Thread(target=_b, daemon=True).start()


# ANSI 控制台颜色
class C:
    R   = "\033[0m"
    RED = "\033[91m"
    GRN = "\033[92m"
    YLW = "\033[93m"
    BLU = "\033[94m"
    CYN = "\033[96m"
    B   = "\033[1m"
    D   = "\033[2m"


# ═══════════════════════════════════════════════════════════════
# 1. Config — 全局配置数据类
# ═══════════════════════════════════════════════════════════════

@dataclass
class Config:
    trigger_key: str = "f9"

    # ── Module A: RTSS 帧率锁定 ──
    rtss_on:       bool   = True
    rtss_hotkey:   str    = "scroll lock"    # RTSS 设置中的切换热键
    rtss_restore:  bool   = True             # 挑战结束后恢复帧率

    # ── Module B: 后台切换 ──
    focus_on:      bool   = False
    focus_method:  str    = "minimize"       # "minimize" | "alt_tab"
    focus_delay:   float  = 0.5              # 检测到黑屏后, 延迟多久切走
    focus_dur:     float  = 2.5              # 后台停留时长 (核心参数)

    # ── Module C: CPU 压力 ──
    cpu_on:        bool   = False
    cpu_cores:     int    = 0                # 0 = 自动 (取一半核心数)

    # ── Module D: 黑屏检测 ──
    sd_on:         bool   = False
    sd_thresh:     int    = 20               # 像素亮度阈值 (<此值视为黑)
    sd_ratio:      float  = 0.70             # 黑色像素占比阈值
    sd_interval:   float  = 0.05             # 采样间隔

    # ── 固定延迟后备 (黑屏检测不可用时) ──
    fb_black:      float  = 0.4              # 触发后到黑屏的估计延迟
    fb_post:       float  = 4.5              # 切回后等321结束的估计时间


# ═══════════════════════════════════════════════════════════════
# 2. Module Base — 模块基类 (6个生命周期钩子)
# ═══════════════════════════════════════════════════════════════

class Mod:
    """
    所有模块的基类.
    定义了编排器在每个阶段会调用的钩子方法,
    子类只需覆盖自己关心的阶段即可.
    """
    name = "base"
    on   = False

    def __init__(self, on=False):
        self.on = on

    def log(self, s):
        ts = time.strftime("%H:%M:%S")
        print(f"  [{ts}] {C.CYN}{self.name.upper()}{C.R} {s}")

    # ── 生命周期钩子 (子类按需覆盖) ──
    def pre(self):            pass    # Phase 0:  挑战前预处理
    def trigger(self, hw):    pass    # 触发瞬间 (传入游戏窗口句柄)
    def black(self):          pass    # 检测到黑屏
    def bg_enter(self):       pass    # 进入后台
    def bg_exit(self):        pass    # 离开后台
    def post(self):           pass    # Phase 5:  挑战后处理
    def clean(self):          pass    # 始终最后调用 (清理资源)


# ═══════════════════════════════════════════════════════════════
# 3. Module A — RTSS Frame Lock
# ═══════════════════════════════════════════════════════════════

class RTSSMod(Mod):
    """
    RTSS 帧率锁定模块
    ─────────────────
    原理:
      高刷显示器(144/240Hz)下, Unity Time.deltaTime 每帧更小,
      加载期间的 Update() 调用频率更高 → 底层计时器每帧消耗的时间更少
      → 看起来"加载没花时间", 但 60s 的 UI 动画是固定步长的 → 产生跳秒.
      锁定 60fps 后, 每帧步进一致, 加载期间不额外"偷步".

    操作:
      模拟按下用户在 RTSS 中配置的全局切换热键 (如 Scroll Lock).
    """
    name = "rtss_lock"

    def __init__(self, on, hotkey, restore):
        super().__init__(on)
        self.hk  = hotkey
        self.rst = restore
        self._lk = False            # 当前是否处于锁定状态

    def pre(self):
        """Phase 0: 锁定帧率"""
        if not self.on:
            return
        self.log(f"发送 RTSS 切换键 [{self.hk}] → 锁定帧率")
        try:
            import keyboard
            keyboard.send(self.hk)
            self._lk = True
            time.sleep(0.15)       # 等待 RTSS 响应
            self.log(f"{C.GRN}✓ 帧率已锁定{C.R}")
        except Exception as e:
            self.log(f"{C.RED}✗ RTSS 控制失败: {e}{C.R}")

    def post(self):
        """Phase 5: 恢复帧率"""
        if not self.on or not self.rst or not self._lk:
            return
        self.log(f"发送 RTSS 切换键 [{self.hk}] → 恢复帧率")
        try:
            import keyboard
            keyboard.send(self.hk)
            self._lk = False
            self.log(f"{C.GRN}✓ 帧率已恢复{C.R}")
        except Exception as e:
            self.log(f"{C.RED}✗ RTSS 恢复失败: {e}{C.R}")


# ═══════════════════════════════════════════════════════════════
# 4. Module B — Screen Detection
# ═══════════════════════════════════════════════════════════════

class ScreenMod(Mod):
    """
    黑屏检测模块
    ─────────────
    用 mss 库截取屏幕中心 200×200 区域,
    统计暗像素占比, 判断是否处于黑屏加载状态.
    
    为后续的"后台切换"提供精确的时机信号,
    避免过早切换 (挑战还没开始) 或过晚切换 (已经321了).
    """
    name = "screen"

    def __init__(self, on, thresh, ratio, interval):
        super().__init__(on)
        self.th   = thresh           # 像素亮度阈值
        self.ratio = ratio           # 黑色占比阈值
        self.itv   = interval        # 采样间隔
        self._ok   = False           # mss 是否可用

    def _chk(self):
        """延迟检查 mss 是否已安装"""
        if not self._ok:
            try:
                import mss
                self._ok = True
            except ImportError:
                self.on = False
                self.log(f"{C.YLW}mss 未安装, 将使用固定延迟{C.R}")
        return self._ok

    def _analyze(self):
        """
        截取屏幕中心 200×200 区域, 返回黑色像素占比.
        原始数据格式: BGRA (4 bytes/pixel)
        采样策略: 每4个像素取1个 (stride=16字节), 只检查B通道.
        对于 200×200 区域, 总共检查 10000 个采样点.
        """
        import mss
        with mss.mss() as sct:
            m = sct.monitors[1]               # 主显示器
            sz = 200
            left = m["left"] + (m["width"]  - sz) // 2
            top  = m["top"]  + (m["height"] - sz) // 2
            raw = sct.grab({
                "left": left, "top": top,
                "width": sz, "height": sz
            }).raw
            bc = tot = 0
            for i in range(0, len(raw), 16):
                if raw[i] < self.th:
                    bc += 1
                tot += 1
            return bc / max(tot, 1)

    def get_ratio(self):
        """获取当前黑色像素占比 (用于校准)"""
        if not self._chk():
            return 0.0
        try:
            return self._analyze()
        except Exception:
            return 0.0

    def wait_black(self, timeout=4.0):
        """
        阻塞等待黑屏出现.
        返回 True=检测到黑屏, False=超时 (编排器将使用固定延迟后备).
        """
        if not self.on or not self._chk():
            return False
        self.log("监控屏幕, 等待黑屏...")
        t0 = time.time()
        while time.time() - t0 < timeout and not _shutdown.is_set():
            r = self._analyze()
            bl = int(r * 30)
            bar = "\u2588" * bl + "\u2591" * (30 - bl)
            sys.stdout.write(f"\r  [{bar}] {r:.0%}  ")
            sys.stdout.flush()
            if r >= self.ratio:
                print()
                self.log(f"{C.GRN}✓ 检测到黑屏 (占比 {r:.0%}){C.R}")
                return True
            time.sleep(self.itv)
        print()
        self.log(f"{C.YLW}等待黑屏超时{C.R}")
        return False

    def calibrate(self):
        """校准模式: 实时显示黑色像素占比, 帮助用户调整阈值"""
        if not self._chk():
            print("  mss 未安装, 无法校准")
            return
        try:
            import keyboard
        except ImportError:
            return
        print(f"\n  {C.B}黑屏校准模式{C.R}")
        print(f"  {C.D}将游戏画面切到需要检测的状态, 按任意键退出{C.R}\n")
        done = threading.Event()
        kb = None
        try:
            kb = keyboard.hook(lambda e: done.set())
            while not done.is_set():
                r = self.get_ratio()
                bl = int(r * 40)
                bar = "\u2588" * bl + "\u2591" * (40 - bl)
                if r >= self.ratio:
                    st = f"{C.GRN}← 判定为黑屏 (阈值 {self.ratio:.0%}){C.R}"
                else:
                    st = f"{C.RED}← 判定为非黑屏{C.R}"
                sys.stdout.write(f"\r  [{bar}] {r:.1%} {st}     ")
                sys.stdout.flush()
                time.sleep(0.1)
        finally:
            if kb:
                try:
                    keyboard.unhook(kb)
                except Exception:
                    pass
        print("\n")


# ═══════════════════════════════════════════════════════════════
# 5. Module C — Focus Switch
# ═══════════════════════════════════════════════════════════════

class FocusMod(Mod):
    """
    后台切换模块
    ─────────────
    原理:
      Windows 在窗口失焦时, 会大幅节流该窗口进程的 CPU 调度.
      Unity 的主线程 Update() 频率从 240fps 骤降到个位数.
      这推迟了 "场景加载完成 → 客户端发送 Ready 包" 的时间点.
      服务端在收到 Ready 后才开始计时, 所以最终:
        EndTimeStamp = ServerTime(received_later) + 60s
        → 客户端显示的初始剩余时间 = EndTimeStamp - ServerNow > 60s
        → 出现 61s / 62s 的"红利时间".

    方法:
            minimize: 仅最小化游戏窗口 (不影响其他屏幕显示)
      alt_tab:  模拟 Alt+Tab (兼容性好, 但部分窗口模式下可能不彻底)
    """
    name = "focus"

    def __init__(self, on, method, delay, dur):
        super().__init__(on)
        self.meth  = method
        self.delay = delay
        self.dur   = dur
        self._hw   = 0              # 游戏窗口句柄

    def trigger(self, hw):
        """记录游戏窗口句柄 (触发时游戏应该在前台)"""
        self._hw = hw

    def black(self):
        """黑屏后延迟一小段再切走 (让黑屏完全建立)"""
        if not self.on:
            return
        self.log(f"延迟 {self.delay}s 后切到后台...")
        time.sleep(self.delay)

    def bg_enter(self):
        """执行后台切换"""
        if not self.on:
            return
        if self.meth == "minimize" and self._hw:
            self.log(f"最小化游戏窗口 (hwnd={self._hw:#x})")
            win_minimize(self._hw)
            # 仅最小化游戏窗口, 避免 Win+D 影响其他屏幕内容
            time.sleep(0.1)
        else:
            # Alt+Tab 备选方案
            try:
                import keyboard
                keyboard.press("alt")
                time.sleep(0.05)
                keyboard.press("tab")
                time.sleep(0.05)
                keyboard.release("tab")
                time.sleep(0.05)
                keyboard.release("alt")
                time.sleep(0.2)
            except Exception as e:
                self.log(f"{C.RED}Alt+Tab 失败: {e}{C.R}")
        self.log(f"{C.YLW}→ 游戏已切到后台{C.R}")

    def bg_exit(self):
        """切回游戏"""
        if not self.on:
            return
        if self.meth == "minimize" and self._hw:
            self.log("恢复游戏窗口")
            win_restore(self._hw)
        else:
            try:
                import keyboard
                keyboard.press("alt")
                time.sleep(0.05)
                keyboard.press("tab")
                time.sleep(0.05)
                keyboard.release("tab")
                time.sleep(0.05)
                keyboard.release("alt")
                time.sleep(0.3)
            except Exception:
                pass
        self.log(f"{C.GRN}→ 游戏已切回前台{C.R}")


# ═══════════════════════════════════════════════════════════════
# 6. Module D — CPU Pressure
# ═══════════════════════════════════════════════════════════════

def _burn(ev):
    """
    CPU 密集空循环 (在独立进程中运行).
    每个进程占满一个物理核心.
    """
    while not ev.is_set():
        pass


class CPUMod(Mod):
    """
    CPU 压力模块
    ─────────────
    原理:
      Unity 的 SceneManager.LoadSceneAsync() 使用后台线程加载 AB 包.
      当 CPU 核心被密集占用时, OS 调度器会降低加载线程的优先级,
      LoadSceneAsync 的 allowSceneActivation 进展变慢.
      配合后台切换, 进一步放大 "客户端 Ready 延迟" 的效果.

    实现:
      用 multiprocessing.Process 创建独立进程 (绕过 GIL),
      每个进程一个空循环占满一个核心.
      后台切换结束时, 通过 Event 信号立即释放所有进程.
    """
    name = "cpu"

    def __init__(self, on, cores):
        super().__init__(on)
        self.cores = cores
        self._ev   = None
        self._ps   = []

    @property
    def _n(self):
        """实际使用的核心数"""
        if self.cores > 0:
            return min(self.cores, os.cpu_count() or 4)
        return max(1, (os.cpu_count() or 4) // 2)

    def bg_enter(self):
        """进后台时启动 CPU 压力"""
        if not self.on:
            return
        n = self._n
        self.log(f"启动 CPU 压力 ({n} 核心)...")
        self._ev = MpEvent()
        for _ in range(n):
            p = Process(target=_burn, args=(self._ev,), daemon=True)
            p.start()
            self._ps.append(p)
        time.sleep(0.1)
        active = sum(1 for p in self._ps if p.is_alive())
        self.log(f"{C.YLW}→ {active}/{n} 进程运行中{C.R}")

    def bg_exit(self):
        """出后台时释放 CPU 压力"""
        if not self.on:
            return
        self._kill()

    def _kill(self):
        if self._ev:
            self._ev.set()
        for p in self._ps:
            if p.is_alive():
                p.join(timeout=1.0)
                if p.is_alive():
                    p.terminate()
        self.log(f"{C.GRN}→ CPU 压力已释放{C.R}")
        self._ps.clear()
        self._ev = None

    def clean(self):
        self._kill()


# ═══════════════════════════════════════════════════════════════
# 7. Orchestrator — 编排器 (核心6阶段时序流水线)
# ═══════════════════════════════════════════════════════════════

class Orch:
    """
    挑战计时优化编排器.
    ────────────────────
    将4个模块按精确时序编排为6阶段流水线:

      Phase 0   预处理        → [A]RTSS 锁帧
      Phase 1   等待黑屏      → [B]屏幕检测 或 固定延迟
      Phase 1.5 黑屏后延迟    → [C]focus.black() (延迟切走)
      Phase 2   进入后台      → [C]最小化/Alt+Tab + [D]CPU压力
      Phase 3   后台等待      → 倒计时 (focus_dur 秒)
      Phase 4   返回游戏      → [D]释放CPU + [C]恢复窗口
      Phase 5   后处理        → 等待321结束 + [A]RTSS恢复
    """

    def __init__(self, cfg: Config):
        self.c    = cfg
        self._run = False
        self._hw  = 0

        # 实例化4个模块
        self.A = RTSSMod(cfg.rtss_on, cfg.rtss_hotkey, cfg.rtss_restore)
        self.B = ScreenMod(cfg.sd_on, cfg.sd_thresh, cfg.sd_ratio, cfg.sd_interval)
        self.C = FocusMod(cfg.focus_on, cfg.focus_method, cfg.focus_delay, cfg.focus_dur)
        self.D = CPUMod(cfg.cpu_on, cfg.cpu_cores)
        self.mods = [self.A, self.B, self.C, self.D]

    def go(self):
        """用户按下触发热键时调用"""
        if self._run:
            print(f"  {C.YLW}⚠ 序列正在运行中, 请等待完成{C.R}")
            return
        self._run = True
        self._hw = get_fg_hwnd()
        for m in self.mods:
            m.trigger(self._hw)
        threading.Thread(target=self._seq, daemon=True).start()

    def _seq(self):
        """核心时序流水线 (在独立线程中运行, 不阻塞热键监听)"""
        c = self.c
        try:
            # ── Phase 0: 预处理 ──
            self._ph("Phase 0", "预处理")
            self.A.pre()                    # RTSS 锁帧
            beep(800, 60)

            # ── Phase 1: 等待黑屏 ──
            self._ph("Phase 1", "等待黑屏")
            ok = self.B.wait_black(4.0) if self.B.on else False
            if not ok:
                self._lg(f"未检测到黑屏, 使用固定延迟 {c.fb_black}s")
                time.sleep(c.fb_black)
            beep(600, 60)

            # ── Phase 1.5: 黑屏后延迟 ──
            self.C.black()                  # focus 模块延迟切走

            # ── Phase 2: 进入后台 ──
            self._ph("Phase 2", "进入后台")
            self.C.bg_enter()               # 最小化/Alt+Tab
            self.D.bg_enter()               # CPU 压力

            # ── Phase 3: 后台等待 ──
            self._ph("Phase 3", f"后台等待 ({c.focus_dur}s)")
            self._cd(c.focus_dur)
            beep(1000, 80)

            # ── Phase 4: 返回游戏 ──
            self._ph("Phase 4", "返回游戏")
            self.D.bg_exit()                # 释放 CPU
            self.C.bg_exit()                # 恢复窗口

            # ── Phase 5: 后处理 ──
            self._ph("Phase 5", f"等待321结束 ({c.fb_post}s)")
            self._cd(c.fb_post, "倒计时")
            self.A.post()                   # RTSS 恢复

            # ── 完成 ──
            print(f"\n  {C.GRN}{C.B}{'='*46}{C.R}")
            print(f"  {C.GRN}{C.B}  ✓ 序列执行完成! 祝挑战顺利!{C.R}")
            print(f"  {C.D}  可再次按触发热键重复执行, 无需重启脚本{C.R}")
            print(f"  {C.GRN}{C.B}{'='*46}{C.R}\n")
            beep(1200, 150)

        except Exception as e:
            self._lg(f"{C.RED}✗ 序列执行出错: {e}{C.R}")
            import traceback
            traceback.print_exc()
        finally:
            for m in self.mods:
                m.clean()
            self._run = False

    def _ph(self, name, desc):
        """打印阶段标题"""
        print(f"\n  {C.B}{C.BLU}── {name}: {desc} ──{C.R}  "
              f"[{time.strftime('%H:%M:%S')}]")

    def _cd(self, secs, label="等待"):
        """带进度条的倒计时"""
        t = int(secs)
        r = secs - t
        for i in range(t, 0, -1):
            if _shutdown.is_set():
                return
            f = int((1 - i / secs) * 20)
            bar = "\u2588" * f + "\u2591" * (20 - f)
            sys.stdout.write(
                f"\r  {C.D}{label}: {bar} {i}s{C.R}  "
            )
            sys.stdout.flush()
            time.sleep(1)
        if r > 0 and not _shutdown.is_set():
            time.sleep(r)
        full_bar = "\u2588" * 20
        sys.stdout.write(
            f"\r  {label}: {full_bar} 完成    \n"
        )
        sys.stdout.flush()

    def _lg(self, s):
        print(f"  [{time.strftime('%H:%M:%S')}] {s}")


# ═══════════════════════════════════════════════════════════════
# 8. CLI — 交互式命令行界面
# ═══════════════════════════════════════════════════════════════

def banner():
    print(f"""{C.CYN}{'='*52}
  原神挑战计时优化器  v1.0
  Genshin Challenge Timer Optimizer
{'='*52}{C.R}

{C.D}基于 Unity 底层引擎行为分析{C.R}
{C.D}利用 RTSS帧率锁定 / 后台切换 / CPU压力 优化计时{C.R}""")


def show_cfg(c):
    """展示当前配置 (避免终端宽字符导致表格错位)"""
    def st(v):
        return f"{C.GRN}ON {C.R}" if v else f"{C.RED}OFF{C.R}"

    core_text = "自动" if c.cpu_cores == 0 else str(c.cpu_cores)
    print(f"""
{C.B}当前配置:{C.R}
    触发热键: {C.YLW}{c.trigger_key}{C.R}
    [A] RTSS帧率锁定: {st(c.rtss_on)}  热键: {c.rtss_hotkey}
    [B] 后台切换:     {st(c.focus_on)}  方式: {c.focus_method}
    [C] CPU压力:      {st(c.cpu_on)}  核心: {core_text}
    [D] 黑屏检测:     {st(c.sd_on)}

    后台停留时长: {c.focus_dur:.1f}s
    黑屏后延迟:   {c.focus_delay:.1f}s
    黑屏阈值:     亮度<{c.sd_thresh} 且占比>{c.sd_ratio:.0%}
""")


def select_preset():
    """
    预设模式选择:
      1. 稳定60s  — 仅RTSS锁帧, 消除高刷跳秒
      2. 追求61-62s — RTSS + 后台切换 + CPU压力 + 黑屏检测
      3. 极限拉满   — 全部启用 + 最长后台停留
      4. 自定义     — 手动开关每个模块
    """
    P = {
        "1": ("稳定60s (推荐)",
              "仅帧率锁定, 消除高刷跳秒, 获得完整60s挑战时间",
              Config(rtss_on=True,
                     focus_on=False, cpu_on=False, sd_on=False)),
        "2": ("追求61-62s (进阶)",
              "帧率锁定+后台切换+CPU压力+黑屏检测, 利用延迟补偿",
              Config(rtss_on=True,
                     focus_on=True, cpu_on=True, sd_on=True,
                     focus_dur=2.5)),
        "3": ("极限拉满 (实验)",
              "全部启用+最长后台停留, 追求极限时间",
              Config(rtss_on=True,
                     focus_on=True, cpu_on=True, sd_on=True,
                     focus_dur=3.5, focus_delay=0.3)),
    }

    print(f"{C.B}请选择预设模式:{C.R}\n")
    for k, (n, d, _) in P.items():
        print(f"  [{k}] {n}")
        print(f"      {C.D}{d}{C.R}")
    print(f"  [4] 自定义模块 (手动开关每个模块)\n")

    while True:
        ch = input(f"  {C.YLW}选择 (1/2/3/4): {C.R}").strip()
        if ch in P:
            show_cfg(P[ch][2])
            return P[ch][2]
        if ch == "4":
            return customize()
        print(f"  {C.RED}无效选择{C.R}")


def customize():
    """交互式自定义配置"""
    # 自定义模式默认全开, 便于直接微调参数
    c = Config(rtss_on=True, focus_on=True, cpu_on=True, sd_on=True)
    print(f"\n  {C.B}自定义模式{C.R}")
    print(f"  {C.D}命令: a/b/c/d 切换模块 | duration/delay/cores/{C.R}")
    print(f"  {C.D}       method/trigger/hotkey/calibrate | 空行确认{C.R}")
    while True:
        show_cfg(c)
        i = input(f"  {C.YLW}> {C.R}").strip().lower()
        if not i:
            break
        elif i == "a":
            c.rtss_on = not c.rtss_on
        elif i == "b":
            c.focus_on = not c.focus_on
        elif i == "c":
            c.cpu_on = not c.cpu_on
        elif i == "d":
            c.sd_on = not c.sd_on
        elif i == "duration":
            try:
                v = input(f"  后台停留时长 (当前 {c.focus_dur}s): ")
                c.focus_dur = float(v) if v.strip() else c.focus_dur
            except ValueError:
                pass
        elif i == "delay":
            try:
                v = input(f"  黑屏后延迟 (当前 {c.focus_delay}s): ")
                c.focus_delay = float(v) if v.strip() else c.focus_delay
            except ValueError:
                pass
        elif i == "cores":
            try:
                v = input("  CPU核心数 (0=自动): ")
                c.cpu_cores = int(v) if v.strip() else c.cpu_cores
            except ValueError:
                pass
        elif i == "method":
            c.focus_method = input(
                f"  切换方式 minimize/alt_tab (当前 {c.focus_method}): "
            ) or c.focus_method
        elif i == "trigger":
            c.trigger_key = input(
                f"  触发热键 (当前 {c.trigger_key}): "
            ) or c.trigger_key
        elif i == "hotkey":
            c.rtss_hotkey = input(
                f"  RTSS切换热键 (当前 {c.rtss_hotkey}): "
            ) or c.rtss_hotkey
        elif i == "calibrate":
            o = Orch(c)
            o.B.calibrate()
        elif i == "help":
            print("""
  可用命令:
    a/b/c/d       切换模块 [A/B/C/D] 开关
    duration      修改后台停留时长 (秒)
    delay         修改黑屏后延迟 (秒)
    cores         修改CPU压力核心数 (0=自动)
    method        切换方式: minimize (推荐) / alt_tab
    trigger       修改触发热键
    hotkey        修改RTSS切换热键
    calibrate     进入黑屏检测校准模式
    help          显示此帮助
    (空行)        确认配置并继续""")
    show_cfg(c)
    return c


# ═══════════════════════════════════════════════════════════════
# 9. Main — 入口
# ═══════════════════════════════════════════════════════════════

def main():
    banner()

    # 管理员权限检查
    if not is_admin():
        print(f"\n  {C.YLW}⚠ 未以管理员身份运行{C.R}")
        print(f"  {C.D}keyboard 库的全局热键需要管理员权限{C.R}")
        print(f"  {C.D}建议: 右键脚本 → 以管理员身份运行{C.R}")
        if input("\n  是否继续? (y/N): ").strip().lower() != "y":
            sys.exit(0)

    # 依赖检查
    miss = []
    try:
        import keyboard
    except ImportError:
        miss.append("keyboard")
    try:
        import mss
    except ImportError:
        pass
    if miss:
        print(f"\n  {C.RED}缺少必要依赖: {', '.join(miss)}{C.R}")
        print(f"  请运行: pip install {' '.join(miss)}")
        if "keyboard" in miss:
            print(f"  可选:   pip install mss  (启用黑屏检测模块)")
        sys.exit(1)

    # mss 可选提示
    mss_ok = False
    try:
        import mss
        mss_ok = True
    except ImportError:
        pass
    if not mss_ok:
        print(f"  {C.YLW}mss 未安装, 黑屏检测模块不可用{C.R}")
        print(f"  {C.D}安装命令: pip install mss{C.R}")

    # 选择模式
    cfg = select_preset()
    orch = Orch(cfg)

    # 信号处理 (Ctrl+C)
    def on_sig(s, f):
        _shutdown.set()
        for m in orch.mods:
            m.clean()
        sys.exit(0)
    signal.signal(signal.SIGINT, on_sig)

    # 注册全局热键
    import keyboard
    keyboard.add_hotkey(cfg.trigger_key, orch.go, suppress=False)

    # 显示启用模块
    act = []
    if cfg.rtss_on:   act.append(f"{C.GRN}[A]RTSS帧率锁定{C.R}")
    if cfg.focus_on:  act.append(f"{C.GRN}[B]后台切换{C.R}")
    if cfg.cpu_on:    act.append(f"{C.GRN}[C]CPU压力{C.R}")
    if cfg.sd_on:     act.append(f"{C.GRN}[D]黑屏检测{C.R}")

    print(f"""
{C.GRN}{C.B}● 准备就绪!{C.R}
{C.D}
  操作流程:
    1. 进入原神, 找到要挑战的副本
    2. 点击 "开始挑战"
    3. 立即按 {C.R}{C.YLW}{C.B}[{cfg.trigger_key}]{C.R}{C.D} 触发优化序列
    4. 脚本自动执行, 蜂鸣提示完成
    5. 享受完整的挑战时间!

  退出: Ctrl+C
{C.R}
  已启用模块: {', '.join(act) if act else C.RED + '无 (请至少启用一个模块)' + C.R}
  {C.B}等待触发 [{cfg.trigger_key}]...{C.R}
""")

    try:
        keyboard.wait()
    except KeyboardInterrupt:
        _shutdown.set()
        print("\n  已退出")


if __name__ == "__main__":
    main()
