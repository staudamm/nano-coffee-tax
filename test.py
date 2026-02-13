import json
import argparse


def create_file(json_string, output_file):
    # Load JSON data
    data = json.loads(json_string)
    text = data[0]['body']
    # Write JSON data to the output file
    with open(output_file, 'w', encoding='utf-8') as file:
        json.dump(text, file, indent=4)
    # print(f"JSON file created: {output_file}")


def main():
    # Create an argument parser
    parser = argparse.ArgumentParser()

    # Add arguments for JSON string and output path
    parser.add_argument("json_string", type=str, help="JSON string input")
    parser.add_argument("excel_path", type=str, help="Path where the Excel files are (template and target)")

    # Parse the arguments
    args = parser.parse_args()
    create_file(args.json_string, args.excel_path)

if __name__ == "__main__":
    main()

