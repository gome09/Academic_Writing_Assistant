# -*- coding: utf-8 -*-
"""构建器抽象基类与注册表（T4-4 插件化）。

每个 Builder 消费一个中间结构 dict（thesis 或 deck 契约），生成一种产物文件。
第三方可通过 register_builder 装饰器注册新构建器（如 HTML/Markdown），
无需修改 main.py 的分发逻辑；--only choices 自动取自注册表。
"""
from __future__ import annotations

from abc import ABC, abstractmethod
from typing import Any

# 注册表：name -> Builder 实例
BUILDERS: dict[str, "Builder"] = {}
# 注册顺序：保证 word 先于 ppt（draft 模式按序生成，日志可读）
_ORDER: list[str] = []


class Builder(ABC):
    """输出构建器抽象基类。

    子类实现 build(data, out_path)，返回输出路径字符串；
    ppt 等构建器可返回 BuildResult（str 子类）以携带 warnings。
    """

    name: str = ""
    ext: str = ""
    label: str = ""      # 显示名（日志/提示用）

    @abstractmethod
    def build(self, data: dict, out_path: str) -> Any:
        """生成产物到 out_path，返回路径（可附 .warnings 属性）。"""
        raise NotImplementedError

    def post_build(self, out_path: Any, args) -> list:
        """产物后处理钩子（如刷新域、导出 PDF）；默认空操作。

        返回告警列表（默认空）。子类可覆盖以实现特定后处理。
        """
        return []


def register_builder(name: str, ext: str, label: str = ""):
    """装饰器：注册 Builder 子类实例到 BUILDERS（T4-4）。

    用法：
        @register_builder("word", ".docx", label="Word")
        class WordBuilder(Builder): ...
    """

    def decorator(cls):
        instance = cls()
        instance.name = name
        instance.ext = ext
        instance.label = label or name
        if name not in BUILDERS:
            _ORDER.append(name)
        BUILDERS[name] = instance
        return cls

    return decorator


def builder_names() -> list[str]:
    """已注册构建器名称，按注册顺序返回（跳过已移除的残留顺序项）。"""
    return [n for n in _ORDER if n in BUILDERS]
