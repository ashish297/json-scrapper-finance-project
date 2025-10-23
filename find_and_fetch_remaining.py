import os
import re
import sys
from typing import List, Set

from request import ensure_output_dir, fetch_and_save_for_perm_id


INPUT_DIR = '/home/ashish/Desktop/json-scrapper/JSON-DATA-ALL'
OUTPUT_DIR = '/home/ashish/Desktop/json-scrapper/remaining-json'


def extract_unknown_permids(input_dir: str) -> List[str]:
    """Scan input_dir for files named like 'unknown-<number>.*' and collect unique <number> tokens.

    Returns a list of unique string IDs in sorted order.
    """
    if not os.path.isdir(input_dir):
        print(f"ERROR: Input directory not found: {input_dir}")
        return []

    pattern = re.compile(r"^unknown-([A-Za-z0-9_-]+)\.")
    unique_ids: Set[str] = set()
    for name in os.listdir(input_dir):
        match = pattern.match(name)
        if match:
            unique_ids.add(match.group(1))

    return sorted(unique_ids)


def main() -> None:
    ids = extract_unknown_permids(INPUT_DIR)
    if not ids:
        print("No unknown-* files found or directory missing.")
        return

    if not ensure_output_dir(OUTPUT_DIR):
        return

    print(f"Found {len(ids)} unique PermIDs to re-fetch.")
    processed = 0
    for perm_id in ids:
        try:
            fetch_and_save_for_perm_id(perm_id=perm_id, output_dir=OUTPUT_DIR)
            processed += 1
        except Exception as e:
            print(f"ERROR handling PermID {perm_id}: {e}")

    print(f"Done. Processed {processed}/{len(ids)} IDs. Output: {OUTPUT_DIR}")


if __name__ == '__main__':
    main()


