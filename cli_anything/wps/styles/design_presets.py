"""设计预设 —— 整合 4 个 Skill 的视觉规范。

来源:
  pptx: 配色方案、母版设计
  scientific-slides: 字体层级、颜色对比度、视觉优先
  pptx-from-layouts: Inner Chapter 模板配色
  slide-excellence: 审查标准
"""

class DesignPreset:
    """统一的设计预设对象。"""

    def __init__(self, name, colors, fonts, spacing, rules):
        self.name = name
        self.colors = colors      # {primary, secondary, accent, dark, light, bg}
        self.fonts = fonts        # {title: (name,size,color), body: (name,size,color), ...}
        self.spacing = spacing    # {margin, gap, card_padding, line_height}
        self.rules = rules        # {"visual_ratio": 0.6, "max_colors": 4, ...}

    def to_dict(self):
        return {"name": self.name, "colors": self.colors, "fonts": self.fonts,
                "spacing": self.spacing, "rules": self.rules}

    def get_rgb_hex(self, color_key: str) -> str:
        """返回指定颜色的十六进制字符串（如 '#1A3C8B'）。

        Args:
            color_key: 颜色键名 —— primary / secondary / accent / dark / light / bg

        Returns:
            十六进制颜色字符串
        """
        rgb = self.colors.get(color_key)
        if not rgb:
            raise KeyError(f"颜色键不存在: {color_key}。可用: {list(self.colors.keys())}")
        return f"#{rgb[0]:02X}{rgb[1]:02X}{rgb[2]:02X}"

    def get_com_color(self, color_key: str) -> int:
        """返回 Windows COM 兼容的 BGR 颜色值（RGBColor 格式）。

        Windows COM 自动化接口（pywin32）使用 BGR 而非 RGB，
        此方法将预设的 RGB 元组转换为 COM 兼容的 OLE_COLOR 整数值。

        Args:
            color_key: 颜色键名

        Returns:
            BGR 整数值（如 0x8B3C1A 对应 RGB(26,60,139)）
        """
        rgb = self.colors.get(color_key)
        if not rgb:
            raise KeyError(f"颜色键不存在: {color_key}。可用: {list(self.colors.keys())}")
        r, g, b = rgb
        return (b << 16) | (g << 8) | r


# ============================================
# 4 大预设主题
# ============================================

PRESETS = {}

# --- 1. 学术风格 (scientific-slides) ---
PRESETS["academic"] = DesignPreset(
    name="学术答辩",
    colors={
        "primary":   (26, 60, 139),     # 深蓝
        "secondary": (230, 119, 51),    # 橙
        "accent":    (24, 128, 80),     # 绿
        "dark":      (34, 34, 34),
        "light":     (245, 248, 252),
        "bg":        (255, 255, 255),
    },
    fonts={
        "title":     ("Arial", 40, (26, 60, 139)),    # 36-44pt 粗体无衬线
        "subtitle":  ("Arial", 22, (100, 100, 100)),
        "body":      ("Arial", 24, (34, 34, 34)),      # 24-28pt 正文
        "caption":   ("Arial", 16, (128, 128, 128)),   # 标注
        "chinese":   ("微软雅黑", 24, (34, 34, 34)),
    },
    spacing={
        "margin": 80,
        "gap": 24,
        "card_padding": 20,
        "line_height": 1.5,
    },
    rules={
        "visual_ratio": 0.65,    # 60-70% 视觉内容
        "max_colors": 5,
        "colorblind_safe": True,
        "one_idea_per_slide": True,
        "white_space": 0.4,      # 40-50% 留白
        "max_bullets": 6,
        "prefer_sans_serif": True,
    }
)

# --- 2. 咨询顾问风格 (pptx-from-layouts) ---
PRESETS["consultant"] = DesignPreset(
    name="咨询顾问",
    colors={
        "primary":   (0, 51, 102),      # 深蓝
        "secondary": (0, 168, 232),     # 亮青
        "accent":    (255, 140, 0),     # 橙
        "dark":      (33, 33, 33),
        "light":     (242, 246, 250),
        "bg":        (255, 255, 255),
    },
    fonts={
        "title":     ("Arial", 36, (0, 51, 102)),
        "subtitle":  ("Arial", 20, (80, 80, 80)),
        "body":      ("Arial", 18, (33, 33, 33)),
        "caption":   ("Arial", 14, (140, 140, 140)),
        "chinese":   ("微软雅黑", 18, (33, 33, 33)),
    },
    spacing={
        "margin": 60,
        "gap": 20,
        "card_padding": 16,
        "line_height": 1.4,
    },
    rules={
        "visual_ratio": 0.5,
        "max_colors": 4,
        "colorblind_safe": False,
        "one_idea_per_slide": True,
        "white_space": 0.4,
        "max_bullets": 5,
        "prefer_sans_serif": True,
        "grid_layout": True,          # 59种网格布局
        "content_density": "medium",
    }
)

# --- 3. 商务汇报风格 (pptx) ---
PRESETS["business"] = DesignPreset(
    name="商务汇报",
    colors={
        "primary":   (0, 82, 148),      # 商务蓝
        "secondary": (200, 40, 40),     # 强调红
        "accent":    (45, 160, 80),     # 绿
        "dark":      (45, 45, 48),
        "light":     (248, 249, 250),
        "bg":        (255, 255, 255),
    },
    fonts={
        "title":     ("Arial", 36, (0, 82, 148)),
        "subtitle":  ("Arial", 20, (90, 90, 90)),
        "body":      ("Arial", 18, (45, 45, 48)),
        "caption":   ("Arial", 14, (140, 140, 140)),
        "chinese":   ("微软雅黑", 18, (45, 45, 48)),
    },
    spacing={
        "margin": 70,
        "gap": 22,
        "card_padding": 18,
        "line_height": 1.4,
    },
    rules={
        "visual_ratio": 0.45,
        "max_colors": 4,
        "colorblind_safe": False,
        "one_idea_per_slide": True,
        "white_space": 0.35,
        "max_bullets": 6,
        "prefer_sans_serif": True,
        "content_density": "medium",
    }
)

# --- 4. 科技/极简风格 ---
PRESETS["tech"] = DesignPreset(
    name="科技极简",
    colors={
        "primary":   (15, 20, 35),       # 近黑
        "secondary": (0, 200, 255),      # 亮青
        "accent":    (255, 100, 60),     # 橙红
        "dark":      (15, 20, 35),
        "light":     (240, 242, 245),
        "bg":        (15, 20, 35),
    },
    fonts={
        "title":     ("Arial", 44, (255, 255, 255)),
        "subtitle":  ("Arial", 20, (160, 180, 200)),
        "body":      ("Arial", 20, (255, 255, 255)),
        "caption":   ("Arial", 14, (120, 130, 150)),
        "chinese":   ("微软雅黑", 20, (255, 255, 255)),
    },
    spacing={
        "margin": 100,
        "gap": 30,
        "card_padding": 24,
        "line_height": 1.6,
    },
    rules={
        "visual_ratio": 0.55,
        "max_colors": 3,
        "colorblind_safe": True,
        "one_idea_per_slide": True,
        "white_space": 0.5,
        "max_bullets": 4,
        "prefer_sans_serif": True,
        "dark_mode": True,
        "content_density": "low",
    }
)
