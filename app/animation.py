# -*- coding: utf-8 -*-
"""通用 UI 动画工具（性能优化版）。

提供：控件淡入淡出、标签页关闭淡出、对话框弹出、按钮悬停背景色过渡等动画。
全部基于 QPropertyAnimation / QVariantAnimation，运行在 Qt 事件循环内，
不占用额外线程、不阻塞 UI 主线程。

针对“动画一顿一顿”的卡顿根因优化：
1. 【复用 QGraphicsOpacityEffect】旧实现每次动画都 new 一个
   QGraphicsOpacityEffect 并 setGraphicsEffect()，每次挂载/卸载都会触发
   widget 整棵子树的 style 重算与重绘风暴。本实现把 effect 登记在 widget
   实例属性（_anim_effect）上，淡入/淡出/弹出统一复用，动画结束后保留
   opacity=1.0 的已挂载 effect（不反复卸载），彻底消除重复新建与重绘风暴。
2. 【复用 QPropertyAnimation / QVariantAnimation】动画对象同样登记复用：
   新动画启动前 stop 旧动画并重设 duration/start/end/easing 后重新 start，
   既避免对象堆积，也实现“先停旧再启新”，不会新旧动画叠加闪烁。
3. 【统一时长与缓动】淡入 180ms OutCubic、淡出 160ms InCubic、
   弹出 200ms OutCubic、悬停 120ms OutCubic，处于 150~250ms 舒适区间，
   流畅不拖沓。
4. 【大面积控件跳过透明度动画】QGraphicsOpacityEffect 会让被作用的控件
   整体渲染到离屏缓冲再合成，面积巨大（如 Excel 表格 / Word 页面）时开销
   明显。超过面积阈值时淡入/淡出直接跳过透明度动画（淡出直接执行回调），
   避免滚动表格出现明显掉帧。
5. 【性能计数】_PERF_STATS 记录 effect / animation 的创建与复用次数，
   供离屏冒烟测试对比优化效果。

生命周期约定（QGraphicsOpacityEffect 防悬垂核心规则，沿用旧版经验）：
- 动画对象以 widget 为父对象，随 widget 销毁；一律不依赖 DeleteWhenStopped
  （该策略下运行中的动画随父析构会二次删除导致原生崩溃）。
- 访问任何 effect / 动画前先用 shiboken6.isValid 判活；
- 卸载特效只用 widget.setGraphicsEffect(None)（Qt 同步删除 C++ 对象），
  严禁再对已卸载的 effect 调用 deleteLater()。
"""
import shiboken6
from PySide6.QtCore import QPropertyAnimation, QVariantAnimation, QEasingCurve, QPoint
from PySide6.QtGui import QColor
from PySide6.QtWidgets import QGraphicsOpacityEffect

#: 统一动画时长（毫秒，处于 150~250ms 舒适区间）
FADE_IN_MS = 180      # 淡入
FADE_OUT_MS = 160     # 淡出
POP_IN_MS = 200       # 对话框弹出
HOVER_MS = 120        # 悬停过渡

#: widget 上登记的动画 / 特效 / 背景动画实例属性名
_ATTR_ANIM = "_anim_opacity"
_ATTR_EFFECT = "_anim_effect"
_ATTR_BG = "_bg_anim"
_ATTR_FIN_CONN = "_anim_finished_connected"
_ATTR_BG_CONN = "_bg_anim_connected"
_ATTR_BG_FIN_CONN = "_bg_anim_finished_connected"
_ATTR_FADING_OUT = "_anim_fading_out"

#: 超过该面积（像素²）的控件跳过透明度动画，避免离屏合成掉帧
_LARGE_WIDGET_AREA = 3_600_000

#: 性能计数（冒烟测试可读取对比优化前后）
_PERF_STATS = {"effects_created": 0, "anims_created": 0,
               "anims_reused": 0, "skipped_large": 0}


def _live(widget, attr):
    """读取 widget 上登记的动画/特效，C++ 对象已失效时返回 None。"""
    obj = getattr(widget, attr, None)
    if obj is not None and shiboken6.isValid(obj):
        return obj
    return None


def _is_large(widget):
    """面积超过阈值的控件判定（Excel 表格 / Word 页面等大滚动控件）。"""
    w, h = widget.width(), widget.height()
    if w * h > _LARGE_WIDGET_AREA:
        _PERF_STATS["skipped_large"] += 1
        return True
    return False


def _effect_for(widget):
    """获取 widget 已挂载的透明度特效；没有则创建并挂载（复用优先）。"""
    effect = _live(widget, _ATTR_EFFECT)
    if effect is None or widget.graphicsEffect() is not effect:
        effect = QGraphicsOpacityEffect(widget)
        widget.setGraphicsEffect(effect)
        setattr(widget, _ATTR_EFFECT, effect)
        _PERF_STATS["effects_created"] += 1
    return effect


def _detach_effects(widget):
    """彻底移除 widget 上的特效（仅淡出结束后需要：widget 即将销毁）。"""
    effect = _live(widget, _ATTR_EFFECT)
    if effect is not None and widget.graphicsEffect() is effect:
        setattr(widget, _ATTR_EFFECT, None)
        widget.setGraphicsEffect(None)  # Qt 同步删除 effect 的 C++ 对象
    setattr(widget, _ATTR_ANIM, None)


def _reset_opacity_anim(widget, effect, duration, start, end):
    """返回可复用的 QPropertyAnimation（复用旧对象，先停旧再启新）。

    复用策略：同一个 widget 的透明度动画对象只创建一次，后续调用
    stop() 后重设参数再 start()，避免对象堆积与新旧动画叠加。
    """
    anim = _live(widget, _ATTR_ANIM)
    if anim is None:
        anim = QPropertyAnimation(effect, b"opacity", widget)
        setattr(widget, _ATTR_ANIM, anim)
        _PERF_STATS["anims_created"] += 1
    else:
        _PERF_STATS["anims_reused"] += 1
    try:
        anim.stop()
    except RuntimeError:
        pass
    # 只有我们注册过 finished 回调才断开，避免无连接时打印 RuntimeWarning
    if getattr(widget, _ATTR_FIN_CONN, False):
        try:
            anim.finished.disconnect()
        except RuntimeError:
            pass
        setattr(widget, _ATTR_FIN_CONN, False)
    anim.setDuration(duration)
    anim.setStartValue(start)
    anim.setEndValue(end)
    anim.setEasingCurve(QEasingCurve.Type.OutCubic)
    return anim


def _mark_finished_conn(widget, anim, slot):
    """为复用动画注册 finished 回调并记录连接标志（去重）。"""
    if not getattr(widget, _ATTR_FIN_CONN, False):
        anim.finished.connect(slot)
        setattr(widget, _ATTR_FIN_CONN, True)


def _on_opacity_done(widget, anim):
    """动画正常结束：仅解除动画登记（保留已挂载 effect 供下次复用）。"""
    if getattr(widget, _ATTR_ANIM, None) is anim:
        setattr(widget, _ATTR_ANIM, None)


def fade_in(widget, duration=FADE_IN_MS, start=0.35, end=1.0):
    """控件淡入。大面积控件跳过透明度动画以避免滚动/表格掉帧。"""
    if _is_large(widget):
        return None
    effect = _effect_for(widget)
    effect.setOpacity(start)
    anim = _reset_opacity_anim(widget, effect, duration, start, end)
    _mark_finished_conn(widget, anim,
                        lambda: _on_opacity_done(widget, anim))
    anim.start()   # KeepWhenStopped + 复用，生命周期随父对象
    return anim


def fade_out_then(widget, callback, duration=FADE_OUT_MS):
    """控件淡出，结束后执行 callback（如移除标签页）。"""
    if _is_large(widget):
        callback()   # 大控件直接执行回调，跳过离屏合成
        return None
    effect = _effect_for(widget)
    effect.setOpacity(1.0)
    anim = _reset_opacity_anim(widget, effect, duration, 1.0, 0.0)
    setattr(widget, _ATTR_FADING_OUT, True)

    def _done():
        setattr(widget, _ATTR_FADING_OUT, False)
        _on_opacity_done(widget, anim)
        # widget 即将销毁/卸载：移除特效恢复普通绘制路径
        _detach_effects(widget)
        callback()

    _mark_finished_conn(widget, anim, _done)
    anim.start()
    return anim


def clear_transition(widget):
    """立即终止 widget 上任何进行中的淡入/弹出过渡，并卸载特效。

    用于页面/标签切换：先把旧内容彻底清出（动画停止 + 特效卸载，
    透明度立即回到不透明），再显示/淡入新内容，杜绝新旧文字在同一帧
    同时渲染造成重叠（功能区字体/字号/加粗/斜体叠影即由此类残留导致）。
    """
    if getattr(widget, _ATTR_FADING_OUT, False):
        # 正在淡出销毁的控件交给其自身 finished 回调收尾，避免中途打断
        return
    anim = _live(widget, _ATTR_ANIM)
    if anim is not None:
        try:
            anim.stop()
        except RuntimeError:
            pass
    # 卸载特效：解除离屏合成，透明度恢复 1.0，不再留下任何半透明叠影
    effect = _live(widget, _ATTR_EFFECT)
    if effect is not None and widget.graphicsEffect() is effect:
        setattr(widget, _ATTR_EFFECT, None)
        widget.setGraphicsEffect(None)  # Qt 同步删除 effect 的 C++ 对象
    setattr(widget, _ATTR_ANIM, None)
    setattr(widget, _ATTR_FIN_CONN, False)


def pop_in(widget, duration=POP_IN_MS, rise=10, start_opacity=0.0):
    """对话框 / 面板弹出动画：轻微上移 + 淡入，模拟 WPS 弹层效果。

    透明度部分复用已登记的 effect 与动画对象；位移部分为轻量 pos 动画，
    随父对象生命周期管理。完成后保留 effect（opacity=1.0），避免重挂风暴。
    """
    end_pos = QPoint(widget.pos())
    start_pos = QPoint(widget.pos().x(), widget.pos().y() + rise)
    widget.move(start_pos)

    effect = _effect_for(widget)
    effect.setOpacity(start_opacity)
    opacity = _reset_opacity_anim(widget, effect, duration, start_opacity, 1.0)
    _mark_finished_conn(widget, opacity,
                        lambda: _on_opacity_done(widget, opacity))

    pos_anim = QPropertyAnimation(widget, b"pos", widget)
    pos_anim.setDuration(duration)
    pos_anim.setStartValue(start_pos)
    pos_anim.setEndValue(end_pos)
    pos_anim.setEasingCurve(QEasingCurve.Type.OutCubic)
    pos_anim.finished.connect(
        lambda: pos_anim.deleteLater() if shiboken6.isValid(pos_anim) else None)

    opacity.start()   # 默认 KeepWhenStopped，生命周期随父对象
    pos_anim.start()
    return opacity


def animate_background(widget, from_color, to_color, duration=HOVER_MS,
                       on_changed=None, on_finished=None):
    """背景色平滑过渡（QVariantAnimation 逐帧回调 setStyleSheet）。

    用于按钮悬停等高频场景；on_changed 可接收当前 QColor 以自定义应用方式。
    动画对象登记在 widget._bg_anim 复用，避免高频悬停反复新建动画对象。
    """
    anim = _live(widget, _ATTR_BG)
    if anim is None:
        anim = QVariantAnimation(widget)
        setattr(widget, _ATTR_BG, anim)
        _PERF_STATS["anims_created"] += 1
    else:
        _PERF_STATS["anims_reused"] += 1
    try:
        anim.stop()
    except RuntimeError:
        pass
    # 仅断开我们自己注册的回调，避免无连接时打印 RuntimeWarning
    if getattr(widget, _ATTR_BG_CONN, False):
        try:
            anim.valueChanged.disconnect()
        except RuntimeError:
            pass
        if getattr(widget, _ATTR_BG_FIN_CONN, False):
            try:
                anim.finished.disconnect()
            except RuntimeError:
                pass
            setattr(widget, _ATTR_BG_FIN_CONN, False)
        setattr(widget, _ATTR_BG_CONN, False)
    anim.setDuration(duration)
    anim.setStartValue(QColor(from_color))
    anim.setEndValue(QColor(to_color))
    anim.setEasingCurve(QEasingCurve.Type.OutCubic)

    if on_changed is None:
        def on_changed(color, _w=widget):
            pass
    anim.valueChanged.connect(lambda v: on_changed(v))
    if on_finished is not None:
        anim.finished.connect(on_finished)
        setattr(widget, _ATTR_BG_FIN_CONN, True)
    setattr(widget, _ATTR_BG_CONN, True)
    anim.start()   # 默认 KeepWhenStopped，生命周期随父对象，避免运行中析构二次删除
    return anim
