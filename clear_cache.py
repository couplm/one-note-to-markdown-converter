import argparse
import json
from pathlib import Path

def clear_cache(output_dir: str = "./onenote_output/"):

    output_path = Path(output_dir)
        
    cache_file = output_path / ".conversion_cache.json"

    if cache_file.exists():
        print(f"Clearing cache from {cache_file}")
        cache = {}
        cache_file.write_text(json.dumps(cache, indent=2))
    else:
        print(f"Cache file {cache_file} does not exist")

def main():
    parser = argparse.ArgumentParser(description="Clear cache from OneNote to Markdown converter")
    parser.add_argument("--output", "-o", help="Output directory", default="./onenote_output")

    args = parser.parse_args()

    try:
        cache = clear_cache(args.output)
    except Exception as e:
        print(e)
        exit(1)

if __name__ == "__main__":
    main()
