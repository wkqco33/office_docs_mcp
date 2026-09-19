"""Mermaid flowchart parser and native PowerPoint shape renderer."""

from __future__ import annotations

import re
from collections import deque
from typing import Any

from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_CONNECTOR, MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.xmlchemy import OxmlElement
from pptx.slide import Slide
from pptx.util import Inches, Pt


def parse_mermaid(
    code: str,
) -> tuple[dict[str, dict[str, Any]], list[tuple[str, str]], str]:
    """Parse a subset of Mermaid flowchart code into nodes, edges, and direction.

    Supports:
        - Direction: 'graph TD', 'flowchart LR', etc.
        - Node shapes:
            - Box: `id[label]`
            - Rounded: `id(label)`
            - Diamond: `id{label}`
            - Stadium: `id([label])`
        - Arrows:
            - `-->`, `---`, `==>`, `-.->`, with inline labels `-->|text|`

    Args:
        code: Mermaid flowchart code string.

    Returns:
        Tuple of (nodes_dict, edges_list, direction).
    """
    nodes: dict[str, dict[str, Any]] = {}
    edges: list[tuple[str, str]] = []
    direction = "TD"

    # 1. Detect direction
    dir_match = re.search(r"\b(?:graph|flowchart)\s+(TD|TB|LR|BT|RL)\b", code, re.IGNORECASE)
    if dir_match:
        d = dir_match.group(1).upper()
        direction = "LR" if d in ("LR", "RL") else "TD"

    # Node pattern: ID followed by bracket delimiters
    node_def_pattern = re.compile(
        r"([a-zA-Z0-9_\-]+)\s*(?:"
        r"\(\[(.*?)\]\)|"  # 1: stadium ([label])
        r"\{(.*?)\}|"  # 2: diamond {label}
        r"\[(.*?)\]|"  # 3: rectangle [label]
        r"\((.*?)\)"  # 4: rounded rectangle (label)
        r")"
    )

    lines = code.strip().splitlines()
    for line in lines:
        line = line.strip()
        if not line or line.startswith("%%"):
            continue
        if re.match(r"^(?:graph|flowchart)\b", line, re.IGNORECASE):
            continue

        # Find all explicit node definitions in this line
        for match in node_def_pattern.finditer(line):
            n_id = match.group(1)
            raw_label = match.group(2) or match.group(3) or match.group(4) or match.group(5) or n_id
            # Clean up line breaks and surrounding quotes
            label = (
                raw_label.strip("\"'")
                .replace("<br/>", "\n")
                .replace("<br>", "\n")
                .replace("<br />", "\n")
            )

            shape_type = "rectangle"
            if match.group(2):
                shape_type = "stadium"
            elif match.group(3):
                shape_type = "diamond"
            elif match.group(4):
                shape_type = "rectangle"
            elif match.group(5):
                shape_type = "rounded_rectangle"

            nodes[n_id] = {"id": n_id, "label": label, "shape_type": shape_type}

        # Clean inline edge labels like -->|text| or -- text -->
        clean_line = re.sub(r"-->\|.*?\|", "-->", line)
        clean_line = re.sub(r"--\s*.*?\s*-->", "-->", clean_line)

        # Replace node brackets so we only have IDs and arrows for edge parsing
        id_only_line = node_def_pattern.sub(r"\1", clean_line)

        # Split by arrows (--> or --- or ==> or -.->)
        arrow_pattern = re.compile(r"\s*(?:-->|---|==>|-\.->)\s*")
        parts = arrow_pattern.split(id_only_line)
        if len(parts) > 1:
            clean_parts = [p.strip() for p in parts if p.strip()]
            for i in range(len(clean_parts) - 1):
                u = clean_parts[i]
                v = clean_parts[i + 1]
                # If node wasn't explicitly defined, add default
                if u not in nodes:
                    nodes[u] = {"id": u, "label": u, "shape_type": "rectangle"}
                if v not in nodes:
                    nodes[v] = {"id": v, "label": v, "shape_type": "rectangle"}
                edges.append((u, v))

    return nodes, edges, direction


def compute_flowchart_layout(
    nodes: dict[str, dict[str, Any]],
    edges: list[tuple[str, str]],
    direction: str,
    left: float,
    top: float,
    width: float,
    height: float,
) -> dict[str, dict[str, float]]:
    """Compute 2D coordinates for all nodes using topological rank.

    Args:
        nodes: Dictionary of node metadata.
        edges: List of directed edges (u, v).
        direction: 'TD' (Top-Down) or 'LR' (Left-Right).
        left: Left margin in inches.
        top: Top margin in inches.
        width: Total allocated width in inches.
        height: Total allocated height in inches.

    Returns:
        Mapping of node_id -> {x, y, w, h, cx, cy}.
    """
    if not nodes:
        return {}

    adj: dict[str, list[str]] = {nid: [] for nid in nodes}
    in_degree: dict[str, int] = dict.fromkeys(nodes, 0)
    for u, v in edges:
        if u in adj and v in in_degree:
            adj[u].append(v)
            in_degree[v] += 1

    queue = deque([nid for nid in nodes if in_degree[nid] == 0])
    if not queue:
        queue.append(next(iter(nodes)))

    node_level: dict[str, int] = dict.fromkeys(queue, 0)
    visited: set[str] = set(queue)

    while queue:
        u = queue.popleft()
        curr_lvl = node_level[u]
        for v in adj.get(u, []):
            if v not in node_level or node_level[v] < curr_lvl + 1:
                node_level[v] = curr_lvl + 1
            if v not in visited:
                visited.add(v)
                queue.append(v)

    # Any disconnected/unvisited nodes fall into level 0
    for nid in nodes:
        if nid not in node_level:
            node_level[nid] = 0

    level_groups: dict[int, list[str]] = {}
    for nid, lvl in node_level.items():
        level_groups.setdefault(lvl, []).append(nid)

    sorted_levels = sorted(level_groups.keys())
    num_levels = max(1, len(sorted_levels))

    node_boxes: dict[str, dict[str, float]] = {}

    for l_idx, lvl in enumerate(sorted_levels):
        nids = level_groups[lvl]
        k = max(1, len(nids))

        if direction == "TD":
            box_w = min(2.0, (width / k) * 0.75)
            box_h = min(0.9, (height / num_levels) * 0.6)
            y_center = top + (height / (num_levels + 1)) * (l_idx + 1)
            y = y_center - box_h / 2.0
            for i, nid in enumerate(nids):
                x_center = left + (width / (k + 1)) * (i + 1)
                x = x_center - box_w / 2.0
                node_boxes[nid] = {
                    "x": x,
                    "y": y,
                    "w": box_w,
                    "h": box_h,
                    "cx": x_center,
                    "cy": y_center,
                }
        else:  # LR
            box_w = min(1.8, (width / num_levels) * 0.65)
            box_h = min(0.85, (height / k) * 0.7)
            x_center = left + (width / (num_levels + 1)) * (l_idx + 1)
            x = x_center - box_w / 2.0
            for i, nid in enumerate(nids):
                y_center = top + (height / (k + 1)) * (i + 1)
                y = y_center - box_h / 2.0
                node_boxes[nid] = {
                    "x": x,
                    "y": y,
                    "w": box_w,
                    "h": box_h,
                    "cx": x_center,
                    "cy": y_center,
                }

    return node_boxes


def render_flowchart(
    slide: Slide,
    nodes: dict[str, dict[str, Any]],
    edges: list[tuple[str, str]],
    direction: str,
    left: float = 0.8,
    top: float = 1.6,
    width: float = 8.4,
    height: float = 5.0,
    title: str | None = None,
) -> None:
    """Render nodes and connector arrows onto a PowerPoint slide.

    Args:
        slide: pptx Slide object to draw on.
        nodes: Parsed nodes dictionary.
        edges: Parsed edges list.
        direction: 'TD' or 'LR'.
        left: Left margin in inches.
        top: Top margin in inches.
        width: Width in inches.
        height: Height in inches.
        title: Optional title to place at the top of the slide.
    """
    # 1. Set title if requested
    if title:
        if slide.shapes.title and slide.shapes.title.has_text_frame:
            slide.shapes.title.text = title
        else:
            tb = slide.shapes.add_textbox(Inches(0.8), Inches(0.5), Inches(8.4), Inches(0.8))
            tf = tb.text_frame
            tf.word_wrap = True
            p = tf.paragraphs[0]
            p.text = title
            p.font.size = Pt(20)
            p.font.bold = True
            p.font.color.rgb = RGBColor(15, 23, 42)

    # 2. Compute coordinates
    boxes = compute_flowchart_layout(nodes, edges, direction, left, top, width, height)

    # 3. Create node shapes
    for nid, node_info in nodes.items():
        if nid not in boxes:
            continue
        box = boxes[nid]
        shape_kind = node_info.get("shape_type", "rectangle")

        if shape_kind == "diamond":
            mso_shape = MSO_SHAPE.DIAMOND
            bg_color = RGBColor(254, 243, 199)  # Light amber
            border_color = RGBColor(245, 158, 11)  # Amber-500
        elif shape_kind == "stadium":
            mso_shape = MSO_SHAPE.OVAL
            bg_color = RGBColor(240, 253, 244)  # Light green
            border_color = RGBColor(34, 197, 94)  # Green-500
        else:
            mso_shape = MSO_SHAPE.ROUNDED_RECTANGLE
            bg_color = RGBColor(239, 246, 255)  # Light blue
            border_color = RGBColor(59, 130, 246)  # Blue-500

        shape = slide.shapes.add_shape(
            mso_shape,
            Inches(box["x"]),
            Inches(box["y"]),
            Inches(box["w"]),
            Inches(box["h"]),
        )

        # Style shape
        shape.fill.solid()
        shape.fill.fore_color.rgb = bg_color
        shape.line.color.rgb = border_color
        shape.line.width = Pt(1.5)

        # Style text frame
        tf = shape.text_frame
        tf.word_wrap = True
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        tf.text = node_info.get("label", nid)

        for paragraph in tf.paragraphs:
            paragraph.alignment = PP_ALIGN.CENTER
            for run in paragraph.runs:
                run.font.size = Pt(11)
                run.font.bold = True
                run.font.color.rgb = RGBColor(30, 41, 59)

    # 4. Create connector arrows for edges
    for u, v in edges:
        if u not in boxes or v not in boxes:
            continue
        b_u = boxes[u]
        b_v = boxes[v]

        if direction == "TD":
            begin_x = b_u["cx"]
            begin_y = b_u["y"] + b_u["h"]
            end_x = b_v["cx"]
            end_y = b_v["y"]
        else:  # LR
            begin_x = b_u["x"] + b_u["w"]
            begin_y = b_u["cy"]
            end_x = b_v["x"]
            end_y = b_v["cy"]

        connector = slide.shapes.add_connector(
            MSO_CONNECTOR.STRAIGHT,
            Inches(begin_x),
            Inches(begin_y),
            Inches(end_x),
            Inches(end_y),
        )
        connector.line.width = Pt(1.5)
        connector.line.color.rgb = RGBColor(100, 116, 139)

        # Add arrowhead at tail end
        try:
            ln = connector.line._get_or_add_ln()
            tail_end = OxmlElement("a:tailEnd")
            tail_end.set("type", "triangle")
            tail_end.set("w", "med")
            tail_end.set("len", "med")
            ln.append(tail_end)
        except Exception:
            # Fallback if XML manipulation is unsupported in certain environments
            pass
