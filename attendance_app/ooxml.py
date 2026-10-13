"""Relationship helpers for Office ZIP packages."""

import posixpath
import re
import xml.etree.ElementTree as ET


def rels_path(part):
    folder, name = posixpath.split(part)
    return posixpath.join(folder, "_rels", name + ".rels") if part else "_rels/.rels"


def prune_package(parts):
    """Remove unreachable parts after intentionally removing template slides/images."""
    reachable, pending = {"[Content_Types].xml"}, [""]
    while pending:
        part = pending.pop()
        if part:
            if part in reachable:
                continue
            if part not in parts:
                raise ValueError(f"양식에 손상된 연결이 있습니다: {part}")
            reachable.add(part)
        relationship = rels_path(part)
        if relationship not in parts:
            continue
        reachable.add(relationship)
        for rel in ET.fromstring(parts[relationship]):
            if rel.get("TargetMode") == "External":
                continue
            target = rel.get("Target", "").split("#")[0]
            if not target:
                continue
            target = target.lstrip("/") if target.startswith("/") else posixpath.normpath(posixpath.join(posixpath.dirname(part), target))
            pending.append(target)
    cleaned = {name: content for name, content in parts.items() if name in reachable}
    content_types = parts["[Content_Types].xml"].decode("utf-8")
    content_types = re.sub(r'<Override\b[^>]*/>', lambda m: m[0] if re.search(r'PartName="([^"]+)"', m[0])[1].lstrip("/") in cleaned else "", content_types)
    cleaned["[Content_Types].xml"] = content_types.encode("utf-8")
    return cleaned


