# -*- coding: utf-8 -*-
"""增量读取缓存（T5-1）。

基于输入文件内容哈希缓存读取后的 Document，二次运行时未变文件跳过重读重整。
- 缓存键：(path, hash, ocr, extract_images)——同文件不同读取选项视为不同结果。
- 仅缓存不含 image 块的 Document（避免大字节存储与深拷贝开销）。
- 持久化为 pickle（本地自有文件，版本号失效旧缓存），存放于 .cache/。
- dry-run 不读写缓存。

本模块为可选增强，read_file 在 cache=None 时行为完全不变。
"""
from __future__ import annotations

import hashlib
import os
import pickle

CACHE_VERSION = 1


def file_hash(path: str) -> str:
    """计算文件内容 SHA256（按 64KB 分块）。"""
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(65536), b""):
            h.update(chunk)
    return h.hexdigest()


def _has_image_blocks(doc) -> bool:
    return any(b.get("kind") == "image" for b in doc.get("blocks", []))


class ReadCache:
    """读取结果缓存：path+hash+options -> Document。

    用法：
        cache = ReadCache(".cache/reads.pkl")
        doc = cache.get(path, h, ocr, extract_images)
        if doc is None:
            doc = read_pdf(...); cache.put(path, h, ocr, extract_images, doc)
        cache.save()
    """

    def __init__(self, path: str):
        self.path = path
        self._data: dict = {}
        self.hits = 0  # 命中次数（自上次 save/load 起累计），供调用方统计
        self._load()

    def _load(self):
        if not os.path.exists(self.path):
            return
        try:
            with open(self.path, "rb") as f:
                obj = pickle.load(f)
            if isinstance(obj, dict) and obj.get("version") == CACHE_VERSION:
                self._data = obj.get("data", {}) or {}
            else:
                self._data = {}  # 版本不匹配，丢弃旧缓存
        except Exception:  # noqa: BLE001 —— 损坏的缓存直接弃用
            self._data = {}

    def get(self, path: str, h: str, ocr: bool, extract_images: bool):
        """命中返回 Document（深拷贝，避免下游变更污染缓存），否则 None。"""
        import copy
        key = (path, h, ocr, extract_images)
        doc = self._data.get(key)
        if doc is None:
            return None
        self.hits += 1
        return copy.deepcopy(doc)

    def put(self, path: str, h: str, ocr: bool, extract_images: bool, doc) -> bool:
        """缓存 Document。含 image 块的不缓存（返回 False）。"""
        if _has_image_blocks(doc):
            return False
        self._data[(path, h, ocr, extract_images)] = doc
        return True

    def save(self):
        """持久化到磁盘。"""
        d = os.path.dirname(self.path)
        if d:
            os.makedirs(d, exist_ok=True)
        with open(self.path, "wb") as f:
            pickle.dump({"version": CACHE_VERSION, "data": self._data}, f)

    def __len__(self):
        return len(self._data)
