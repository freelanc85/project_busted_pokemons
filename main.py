"""Scrape the Radical Red Pokédex (https://dex.radicalred.net/) into an Excel sheet.

The dex page can be personalised by uploading a Radical Red save file. This
script drives Chrome with Selenium to do that, scrolls the lazily-loaded species
table until every row is present, caches the resulting HTML, then parses it and
writes one row per species to an .xlsx workbook.

Examples:
    python main.py --save-file "Radical Red v4.1.sav"
    python main.py --pokemon Bulbasaur --pokemon Charizard
    python main.py --refresh --headless --save-file my.sav
"""

from __future__ import annotations

import argparse
import logging
import os
import sys
import tempfile
import time
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable

from bs4 import BeautifulSoup, Tag
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font
from openpyxl.utils import get_column_letter

log = logging.getLogger("rrdex")

DEX_URL = "https://dex.radicalred.net/"
EXCEL_CELL_LIMIT = 32_767  # Max characters Excel allows in a single cell.
STAT_NAMES = ("HP", "Atk", "Def", "SpA", "SpD", "Spe", "BST")

# (header, column width, centred?)
COLUMNS: tuple[tuple[str, float, bool], ...] = (
    ("DexID", 8, True),
    ("Name", 30, False),
    ("Types", 21.5, True),
    ("Primary Ability", 23.5, True),
    ("Secondary Ability", 23.5, True),
    ("Hidden Ability", 23.5, True),
    *((stat, 8, True) for stat in STAT_NAMES),
    ("Sprite (base64)", 20, False),
)


class ScrapeError(RuntimeError):
    """Raised when the dex page cannot be fetched or parsed."""


# --------------------------------------------------------------------------- #
# Parsing
# --------------------------------------------------------------------------- #
@dataclass(frozen=True)
class Pokemon:
    dex_id: int
    name: str
    types: str
    primary_ability: str
    secondary_ability: str
    hidden_ability: str
    stats: tuple[int, ...]
    sprite_base64: str

    def as_row(self, include_sprite: bool) -> list[object]:
        sprite = self.sprite_base64 if include_sprite else ""
        if len(sprite) > EXCEL_CELL_LIMIT:
            log.warning("Sprite for %s exceeds Excel's cell limit; omitted.", self.name)
            sprite = ""
        return [
            self.dex_id,
            self.name,
            self.types,
            self.primary_ability,
            self.secondary_ability,
            self.hidden_ability,
            *self.stats,
            sprite,
        ]


def _text(parent: Tag, name: str, cls: str, default: str = "-") -> str:
    node = parent.find(name, class_=cls)
    text = node.get_text(strip=True) if node else ""
    return text or default


def _row_name(row: Tag) -> str:
    return _text(row, "td", "speciesNameWrapper", default="")


def parse_row(row: Tag) -> Pokemon:
    """Convert one ``tr.speciesRow`` into a :class:`Pokemon`. Raises ValueError if malformed."""
    name = _row_name(row)
    if not name:
        raise ValueError("row has no species name")

    dex_text = _text(row, "td", "speciesDexIDWrapper", default="")
    if not dex_text.isdigit():
        raise ValueError(f"{name}: invalid DexID {dex_text!r}")

    stats = tuple(
        int(s.get_text(strip=True))
        for s in row.find_all("div", class_="speciesStatValue")
        if s.get_text(strip=True).lstrip("-").isdigit()
    )
    if len(stats) != len(STAT_NAMES):
        raise ValueError(f"{name}: expected {len(STAT_NAMES)} stats, found {len(stats)}")

    sprite = ""
    img = row.find("img", class_="speciesSprite")
    src = img.get("src", "") if img else ""
    if isinstance(src, str) and src.startswith("data:") and "," in src:
        sprite = src.split(",", 1)[1]

    types = ", ".join(t.get_text(strip=True) for t in row.find_all("div", class_="typeWrapper"))

    return Pokemon(
        dex_id=int(dex_text),
        name=name,
        types=types or "-",
        primary_ability=_text(row, "div", "speciesAbilitiesPrimary"),
        secondary_ability=_text(row, "div", "speciesAbilitiesSecondary"),
        hidden_ability=_text(row, "div", "speciesAbilitiesHidden"),
        stats=stats,
        sprite_base64=sprite,
    )


def species_rows(html: str) -> list[Tag]:
    return BeautifulSoup(html, "lxml").find_all("tr", class_="speciesRow")


def parse_pokemon(html: str, wanted: set[str]) -> list[Pokemon]:
    """Parse every species row, optionally filtered by lowercase name. Bad rows are skipped."""
    rows = species_rows(html)
    if not rows:
        raise ScrapeError("No species rows found in HTML; the page layout may have changed.")

    result: list[Pokemon] = []
    skipped = 0
    for row in rows:
        if wanted and _row_name(row).lower() not in wanted:
            continue
        try:
            result.append(parse_row(row))
        except (ValueError, TypeError) as exc:
            skipped += 1
            log.warning("Skipping malformed row: %s", exc)

    if wanted:
        missing = wanted - {p.name.lower() for p in result}
        if missing:
            log.warning("Not found in dex: %s", ", ".join(sorted(missing)))
    log.info("Parsed %d Pokémon (%d skipped).", len(result), skipped)
    return result


# --------------------------------------------------------------------------- #
# Excel output
# --------------------------------------------------------------------------- #
def write_workbook(pokemon: Iterable[Pokemon], path: Path, include_sprite: bool) -> int:
    wb = Workbook()
    ws = wb.active
    ws.title = "Pokedex"
    ws.append([header for header, _, _ in COLUMNS])

    count = 0
    for p in pokemon:
        ws.append(p.as_row(include_sprite))
        count += 1

    header_font = Font(bold=True, underline="single")
    centre = Alignment(horizontal="center")
    for idx, (_, width, centred) in enumerate(COLUMNS, start=1):
        letter = get_column_letter(idx)
        ws.column_dimensions[letter].width = width
        ws.cell(row=1, column=idx).font = header_font
        if centred:
            for (cell,) in ws.iter_rows(min_row=1, min_col=idx, max_col=idx):
                cell.alignment = centre

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)
    return count


# --------------------------------------------------------------------------- #
# Fetching
# --------------------------------------------------------------------------- #
def atomic_write_text(path: Path, text: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    fd, tmp = tempfile.mkstemp(dir=path.parent, prefix=f".{path.name}.", suffix=".tmp")
    try:
        with os.fdopen(fd, "w", encoding="utf-8") as fh:
            fh.write(text)
        os.replace(tmp, path)
    except BaseException:
        Path(tmp).unlink(missing_ok=True)
        raise


def load_cached_html(path: Path, wanted: set[str]) -> str | None:
    """Return cached HTML if it exists and contains every wanted Pokémon."""
    if not path.is_file():
        return None
    html = path.read_text(encoding="utf-8")
    names = {_row_name(r).lower() for r in species_rows(html)}
    if not names:
        log.info("Cached HTML at %s has no species rows; refetching.", path)
        return None
    missing = wanted - names
    if missing:
        log.info("Cache is missing %s; refetching.", ", ".join(sorted(missing)))
        return None
    log.info("Using cached HTML from %s (%d species).", path, len(names))
    return html


def fetch_html(args: argparse.Namespace) -> str:
    """Drive Chrome to load the dex (optionally with a save file) and return the page source."""
    # Imported lazily so parsing a cached page doesn't require Selenium.
    from selenium import webdriver
    from selenium.common.exceptions import TimeoutException, WebDriverException
    from selenium.webdriver.chrome.options import Options
    from selenium.webdriver.chrome.service import Service
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support import expected_conditions as EC
    from selenium.webdriver.support.ui import WebDriverWait

    options = Options()
    if args.headless:
        options.add_argument("--headless=new")
    options.add_argument("--window-size=1400,1000")
    if args.chrome_binary:
        options.binary_location = str(args.chrome_binary)
    # Without an explicit driver path, Selenium Manager downloads a matching chromedriver.
    service = Service(str(args.chromedriver)) if args.chromedriver else Service()

    log.info("Starting Chrome%s...", " (headless)" if args.headless else "")
    try:
        driver = webdriver.Chrome(service=service, options=options)
    except WebDriverException as exc:
        raise ScrapeError(f"Could not start Chrome: {exc.msg}") from exc

    try:
        wait = WebDriverWait(driver, args.timeout)
        driver.get(DEX_URL)
        log.info("Opened %s", DEX_URL)

        try:
            wait.until(EC.presence_of_element_located((By.ID, "speciesTable")))
        except TimeoutException as exc:
            raise ScrapeError("Species table never appeared; is the site up?") from exc

        if args.save_file:
            # Send the path straight to the hidden <input type=file>; no OS dialog needed.
            file_input = wait.until(EC.presence_of_element_located((By.ID, "saveFileInput")))
            file_input.send_keys(str(args.save_file.resolve()))
            log.info("Uploaded save file %s", args.save_file)
            time.sleep(1)  # Let the page re-render the table with save data.

        scroll_until_loaded(driver, By, args.timeout)
        return driver.page_source
    finally:
        driver.quit()
        log.info("Closed browser.")


def scroll_until_loaded(driver, by, timeout: float, settle_rounds: int = 5, pause: float = 0.2) -> None:
    """Scroll to the bottom until the species row count stops growing (or timeout)."""
    deadline = time.monotonic() + max(timeout, 5) * 6
    last_count, stable = -1, 0
    while time.monotonic() < deadline:
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        time.sleep(pause)
        count = len(driver.find_elements(by.CSS_SELECTOR, "tr.speciesRow"))
        if count == last_count:
            stable += 1
            if stable >= settle_rounds and count > 0:
                log.info("Table fully loaded: %d species rows.", count)
                return
        else:
            stable = 0
            last_count = count
            log.debug("Loaded %d rows so far...", count)
    if last_count <= 0:
        raise ScrapeError("Timed out before any species rows loaded.")
    log.warning("Scroll timeout reached; continuing with %d rows (may be incomplete).", last_count)


# --------------------------------------------------------------------------- #
# CLI
# --------------------------------------------------------------------------- #
def build_parser() -> argparse.ArgumentParser:
    env = os.environ.get
    p = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument("--save-file", type=Path, default=env("RRDEX_SAVE_FILE"),
                   help="Radical Red .sav file to upload (env: RRDEX_SAVE_FILE). Optional.")
    p.add_argument("-p", "--pokemon", action="append", default=[], metavar="NAME",
                   help="Only export this Pokémon (repeatable). Default: all.")
    p.add_argument("--cache", type=Path, default=Path("html_content.html"),
                   help="Where to cache the fetched HTML (default: %(default)s).")
    p.add_argument("-o", "--output", type=Path, default=Path("pokemon_data.xlsx"),
                   help="Output workbook (default: %(default)s).")
    p.add_argument("--refresh", action="store_true", help="Ignore the cache and refetch.")
    p.add_argument("--delete-cache", action="store_true", help="Delete the HTML cache when done.")
    p.add_argument("--no-sprites", action="store_true", help="Leave the sprite column empty.")
    p.add_argument("--headless", action="store_true", help="Run Chrome without a window.")
    p.add_argument("--chrome-binary", type=Path, default=env("RRDEX_CHROME_BINARY"),
                   help="Chrome executable (env: RRDEX_CHROME_BINARY). Default: auto-detect.")
    p.add_argument("--chromedriver", type=Path, default=env("RRDEX_CHROMEDRIVER"),
                   help="chromedriver path (env: RRDEX_CHROMEDRIVER). Default: Selenium Manager.")
    p.add_argument("--timeout", type=float, default=15, help="Page wait timeout in seconds.")
    p.add_argument("-v", "--verbose", action="store_true", help="Debug logging.")
    return p


def validate_args(args: argparse.Namespace, parser: argparse.ArgumentParser) -> None:
    for attr in ("save_file", "chrome_binary", "chromedriver"):
        value = getattr(args, attr)
        if value is not None:
            value = Path(value).expanduser()
            setattr(args, attr, value)
            if not value.is_file():
                parser.error(f"--{attr.replace('_', '-')}: file not found: {value}")
    if args.timeout <= 0:
        parser.error("--timeout must be positive")


def main(argv: list[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)
    logging.basicConfig(level=logging.DEBUG if args.verbose else logging.INFO,
                        format="%(levelname)s: %(message)s")
    validate_args(args, parser)

    wanted = {name.strip().lower() for name in args.pokemon if name.strip()}

    try:
        html = None if args.refresh else load_cached_html(args.cache, wanted)
        if html is None:
            html = fetch_html(args)
            atomic_write_text(args.cache, html)
            log.info("Cached HTML to %s", args.cache)

        pokemon = parse_pokemon(html, wanted)
        if not pokemon:
            log.error("Nothing to write.")
            return 1
        count = write_workbook(pokemon, args.output, include_sprite=not args.no_sprites)
        log.info("Wrote %d Pokémon to %s", count, args.output)
    except ScrapeError as exc:
        log.error("%s", exc)
        return 1
    except KeyboardInterrupt:
        log.error("Interrupted.")
        return 130

    if args.delete_cache:
        args.cache.unlink(missing_ok=True)
        log.info("Deleted cache %s", args.cache)
    return 0


if __name__ == "__main__":
    sys.exit(main())
