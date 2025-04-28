"""Color palette definitions for modern presentation designs."""
from pptx.dml.color import RGBColor

class ColorPalette:
    """Modern professional color palettes."""
    
    @staticmethod
    def hex_to_rgb(hex_color):
        """Convert hex color to RGB."""
        hex_color = hex_color.lstrip('#')
        return RGBColor(
            int(hex_color[0:2], 16),
            int(hex_color[2:4], 16),
            int(hex_color[4:6], 16)
        )
    
    @staticmethod
    def get_palette(theme="minimalist_blue"):
        """Get color palette by theme name."""
        palettes = {
            "minimalist_blue": {
                "background": ColorPalette.hex_to_rgb("#FFFFFF"),  # White
                "shape": ColorPalette.hex_to_rgb("#E6F0FA"),      # Light Blue
                "text": ColorPalette.hex_to_rgb("#2C3E50"),       # Dark Navy
                "title": ColorPalette.hex_to_rgb("#1A365D")       # Deep Blue
            },
            "soft_gray": {
                "background": ColorPalette.hex_to_rgb("#FFFFFF"),  # White
                "shape": ColorPalette.hex_to_rgb("#F2F2F2"),      # Light Gray
                "text": ColorPalette.hex_to_rgb("#333333"),       # Dark Gray
                "title": ColorPalette.hex_to_rgb("#1A1A1A")       # Almost Black
            },
            "fresh_green": {
                "background": ColorPalette.hex_to_rgb("#FFFFFF"),  # White
                "shape": ColorPalette.hex_to_rgb("#E6F7E6"),      # Light Green
                "text": ColorPalette.hex_to_rgb("#2E8B57"),       # Dark Green
                "title": ColorPalette.hex_to_rgb("#1B5E20")       # Forest Green
            },
            "elegant_purple": {
                "background": ColorPalette.hex_to_rgb("#F8F9FA"),  # Light Gray
                "shape": ColorPalette.hex_to_rgb("#EFE6FA"),      # Soft Lavender
                "text": ColorPalette.hex_to_rgb("#4B0082"),       # Dark Purple
                "title": ColorPalette.hex_to_rgb("#2E1437")       # Deep Purple
            },
            "professional_teal": {
                "background": ColorPalette.hex_to_rgb("#FFFFFF"),  # White
                "shape": ColorPalette.hex_to_rgb("#E6FAF7"),      # Light Teal
                "text": ColorPalette.hex_to_rgb("#008080"),       # Dark Teal
                "title": ColorPalette.hex_to_rgb("#004D4D")       # Deep Teal
            }
        }
        return palettes.get(theme, palettes["minimalist_blue"])
