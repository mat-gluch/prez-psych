## Pixel art asset

Plik `assets/papal_primarch_pixel_art.png` jest generowany deterministycznie przez
`scripts/generate_papal_primarch_pixel_art.py`.

### Regeneracja

```bash
cd /home/runner/work/prez-psych/prez-psych
python -m pip install pillow
python scripts/generate_papal_primarch_pixel_art.py
```

Skrypt rysuje bazę 64×64 i skaluje ją nearest-neighbor do 512×512, aby zachować ostre piksele.
