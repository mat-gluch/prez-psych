from pathlib import Path

from PIL import Image, ImageDraw

BASE_SIZE = 64
SCALE = 8
OUTPUT_PATH = Path(__file__).resolve().parents[1] / "assets" / "papal_primarch_pixel_art.png"

PALETTE = {
    "bg": (8, 10, 20, 255),
    "fog": (28, 31, 52, 255),
    "halo_outer": (255, 212, 90, 255),
    "halo_inner": (255, 242, 176, 255),
    "white": (244, 244, 238, 255),
    "white_shadow": (204, 206, 205, 255),
    "gold": (224, 173, 63, 255),
    "gold_shadow": (171, 118, 38, 255),
    "red": (164, 38, 40, 255),
    "red_shadow": (109, 23, 28, 255),
    "staff": (201, 151, 80, 255),
    "staff_dark": (120, 84, 38, 255),
    "face": (237, 208, 184, 255),
    "line": (34, 24, 22, 255),
    "glow": (255, 245, 215, 255),
}


def draw_centered_ellipse(draw: ImageDraw.ImageDraw, box, color):
    draw.ellipse(box, fill=color)


def main() -> None:
    base = Image.new("RGBA", (BASE_SIZE, BASE_SIZE), PALETTE["bg"])
    draw = ImageDraw.Draw(base)

    # Atmospheric background
    draw.rectangle((0, 44, 63, 63), fill=PALETTE["fog"])
    draw.rectangle((0, 50, 63, 63), fill=(18, 16, 24, 255))
    for x in range(4, 60, 8):
        draw.rectangle((x, 46, x + 2, 49), fill=(45, 43, 58, 255))

    # Energy halo
    draw_centered_ellipse(draw, (16, 1, 48, 23), PALETTE["halo_outer"])
    draw_centered_ellipse(draw, (19, 4, 45, 20), PALETTE["halo_inner"])
    draw_centered_ellipse(draw, (22, 7, 42, 17), PALETTE["bg"])

    # Shoulder silhouette for monumental look
    draw.polygon(
        [(12, 38), (20, 30), (44, 30), (52, 38), (47, 43), (17, 43)],
        fill=PALETTE["gold_shadow"],
    )

    # Red cape
    draw.polygon(
        [(14, 28), (10, 50), (18, 58), (24, 43), (24, 30)],
        fill=PALETTE["red"],
    )
    draw.polygon(
        [(50, 28), (40, 30), (40, 43), (46, 58), (54, 50)],
        fill=PALETTE["red_shadow"],
    )

    # White papal silhouette + armor core
    draw.rectangle((24, 18, 40, 44), fill=PALETTE["white"])
    draw.rectangle((25, 19, 39, 21), fill=PALETTE["white_shadow"])
    draw.rectangle((25, 38, 39, 44), fill=PALETTE["white_shadow"])

    # Head and mitre-inspired helmet
    draw.rectangle((27, 13, 37, 18), fill=PALETTE["white"])
    draw.rectangle((29, 14, 35, 17), fill=PALETTE["face"])
    draw.polygon([(27, 13), (32, 8), (37, 13)], fill=PALETTE["white"])

    # Gold trim and pauldrons
    draw.rectangle((22, 21, 24, 31), fill=PALETTE["gold"])
    draw.rectangle((40, 21, 42, 31), fill=PALETTE["gold"])
    draw.rectangle((24, 22, 40, 24), fill=PALETTE["gold"])
    draw.rectangle((26, 27, 38, 29), fill=PALETTE["gold"])
    draw.rectangle((27, 31, 37, 33), fill=PALETTE["gold_shadow"])
    draw.rectangle((22, 29, 24, 33), fill=PALETTE["gold_shadow"])
    draw.rectangle((40, 29, 42, 33), fill=PALETTE["gold_shadow"])

    # Ornamental crozier (left side)
    draw.rectangle((18, 16, 19, 52), fill=PALETTE["staff"])
    draw.rectangle((19, 16, 20, 52), fill=PALETTE["staff_dark"])
    draw.arc((14, 8, 24, 18), start=210, end=25, fill=PALETTE["staff"], width=2)
    draw.rectangle((16, 11, 17, 13), fill=PALETTE["glow"])

    # Face details and subtle linework
    draw.point((31, 15), fill=PALETTE["line"])
    draw.point((33, 15), fill=PALETTE["line"])
    draw.line((31, 16, 33, 16), fill=PALETTE["line"], width=1)

    # Highlights
    draw.rectangle((28, 23, 30, 24), fill=PALETTE["glow"])
    draw.rectangle((34, 30, 36, 31), fill=PALETTE["glow"])
    draw.rectangle((24, 35, 25, 36), fill=PALETTE["glow"])

    output = base.resize((BASE_SIZE * SCALE, BASE_SIZE * SCALE), Image.Resampling.NEAREST)
    OUTPUT_PATH.parent.mkdir(parents=True, exist_ok=True)
    output.save(OUTPUT_PATH, format="PNG")
    print(f"Saved {OUTPUT_PATH} ({output.width}x{output.height})")


if __name__ == "__main__":
    main()
