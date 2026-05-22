import json
import os
import argparse

def merge_translations(part1_path, part2_path, output_path):
    if not os.path.exists(part1_path):
        print(f"Error: Part 1 file not found: {part1_path}")
        return False
    if not os.path.exists(part2_path):
        print(f"Error: Part 2 file not found: {part2_path}")
        return False

    with open(part1_path, 'r', encoding='utf-8') as f:
        part1 = json.load(f)

    with open(part2_path, 'r', encoding='utf-8') as f:
        part2 = json.load(f)

    all_trans = part1 + part2

    with open(output_path, 'w', encoding='utf-8') as f:
        json.dump(all_trans, f, ensure_ascii=False, indent=2)

    print(f"Merge completed! Combined {len(all_trans)} items into {output_path}")
    return True

if __name__ == "__main__":
    script_dir = os.path.dirname(os.path.abspath(__file__))
    parent_dir = os.path.dirname(script_dir)
    
    parser = argparse.ArgumentParser(description="Merge two translation JSON slices into a single unified JSON.")
    parser.add_argument("--part1", default=os.path.join(parent_dir, "examples", "trans_result_part1.json"),
                        help="Path to translation part 1 JSON")
    parser.add_argument("--part2", default=os.path.join(parent_dir, "examples", "trans_result_part2.json"),
                        help="Path to translation part 2 JSON")
    parser.add_argument("-o", "--output", default=os.path.join(parent_dir, "examples", "sample_result.json"),
                        help="Path to save the combined result JSON")
    
    args = parser.parse_args()
    
    os.makedirs(os.path.dirname(args.output), exist_ok=True)
    merge_translations(args.part1, args.part2, args.output)
