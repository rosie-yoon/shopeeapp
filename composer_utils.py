import io
from pathlib import Path
from PIL import Image, ImageFilter

SHADOW_PRESETS = {
    "off": {"blur": 0, "alpha": 0, "offset_x": 0.0, "offset_y": 0.0},
    "light": {"blur": 6, "alpha": 100, "offset_x": 0.006, "offset_y": 0.006},
    "medium": {"blur": 14, "alpha": 160, "offset_x": 0.012, "offset_y": 0.012},
    "strong": {"blur": 24, "alpha": 220, "offset_x": 0.018, "offset_y": 0.018},
}


def ensure_rgba(img: Image.Image) -> Image.Image:
    if img.mode == "RGBA":
        return img
    if img.mode in ("LA", "P"):
        return img.convert("RGBA")
    if img.mode == "RGB":
        rgba = Image.new("RGBA", img.size, (0, 0, 0, 0))
        rgba.paste(img, (0, 0))
        return rgba
    return img.convert("RGBA")


def has_useful_alpha(img: Image.Image) -> bool:
    img = ensure_rgba(img)
    a = img.getchannel("A")
    extrema = a.getextrema()
    if not extrema:
        return False
    min_a, max_a = extrema
    return not (min_a == 255 and max_a == 255) and not (min_a == 0 and max_a == 0)


def compute_anchor_position(bg_size, fg_size, anchor: str):
    W, H = bg_size
    w, h = fg_size
    positions = {
        "center": ((W - w) // 2, (H - h) // 2),
        "top": ((W - w) // 2, 0),
        "bottom": ((W - w) // 2, H - h),
        "left": (0, (H - h) // 2),
        "right": (W - w, (H - h) // 2),
        "top-left": (0, 0),
        "top-right": (W - w, 0),
        "bottom-left": (0, H - h),
        "bottom-right": (W - w, H - h),
    }
    return positions.get(anchor, positions["center"])


def compose_one_bytes(item_img: Image.Image, template_img: Image.Image, **opts):
    item_rgba = ensure_rgba(item_img)
    template_rgba = ensure_rgba(template_img)

    ratio = float(opts.get("resize_ratio", 1.0))
    if ratio <= 0:
        ratio = 1.0
    if ratio != 1.0:
        new_size = (max(1, int(item_rgba.width * ratio)), max(1, int(item_rgba.height * ratio)))
        item_rgba = item_rgba.resize(new_size, Image.LANCZOS)

    anchor = opts.get("anchor", "center")
    x, y = compute_anchor_position(template_rgba.size, item_rgba.size, anchor)

    composition_mode = opts.get("composition_mode", "normal")
    item_has_alpha = has_useful_alpha(item_rgba)

    if composition_mode == "frame":
        final_img = Image.new("RGBA", template_rgba.size, (255, 255, 255, 255))

        item_layer = Image.new("RGBA", final_img.size, (0, 0, 0, 0))
        item_layer.paste(item_rgba, (x, y), item_rgba)
        final_img = Image.alpha_composite(final_img, item_layer)

        template_layer = Image.new("RGBA", final_img.size, (0, 0, 0, 0))
        template_layer.paste(template_rgba, (0, 0), template_rgba)
        final_img = Image.alpha_composite(final_img, template_layer)
    else:
        final_img = template_rgba.copy()

        if item_has_alpha:
            preset_name = str(opts.get("shadow_preset", "off"))
            preset = SHADOW_PRESETS.get(preset_name, SHADOW_PRESETS["off"])

            if preset.get("alpha", 0) > 0:
                alpha_mask = item_rgba.getchannel("A")

                blur_radius = int(preset.get("blur", 0))
                if blur_radius > 0:
                    alpha_blurred = alpha_mask.filter(ImageFilter.GaussianBlur(blur_radius))
                else:
                    alpha_blurred = alpha_mask

                scale = max(0, min(255, int(preset.get("alpha", 0)))) / 255.0
                alpha_scaled = alpha_blurred.point(lambda p: int(p * scale))

                shadow_rgba = Image.new("RGBA", item_rgba.size, (0, 0, 0, 0))
                shadow_rgba.putalpha(alpha_scaled)

                dx = int(template_rgba.width * float(preset.get("offset_x", 0.0)))
                dy = int(template_rgba.height * float(preset.get("offset_y", 0.0)))

                shadow_layer = Image.new("RGBA", final_img.size, (0, 0, 0, 0))
                shadow_layer.paste(shadow_rgba, (x + dx, y + dy), shadow_rgba)
                final_img = Image.alpha_composite(final_img, shadow_layer)

        item_layer = Image.new("RGBA", final_img.size, (0, 0, 0, 0))
        item_layer.paste(item_rgba, (x, y), item_rgba)
        final_img = Image.alpha_composite(final_img, item_layer)

    img_buf = io.BytesIO()
    out_format = str(opts.get("out_format", "JPEG")).upper()

    if out_format == "JPEG":
        if final_img.mode == 'RGBA':
            background = Image.new("RGB", final_img.size, (255, 255, 255))
            background.paste(final_img, mask=final_img.split()[3])
            final_img = background
        else:
            final_img = final_img.convert("RGB")
        final_img.save(img_buf, format="JPEG", quality=int(opts.get("quality", 92)))
        ext = "jpg"
    else:
        final_img.save(img_buf, format="PNG")
        ext = "png"

    img_buf.seek(0)
    return img_buf, ext
