"""
Export Sigma formatted detection library to an excel file.

twitter: @Cyb3rDefender
"""

from pathlib import Path
import pandas as pd
import yaml
import argparse


class DetectionData:
    def __init__(self, input_dir: Path, output_file: Path, logic: bool) -> None:

        self.detection_paths_list = input_dir.rglob("*.yml")
        self.output_file = output_file
        self.logic = logic

    def fetch(self) -> list:

        # Loop over detection files and add to the detections list
        detections = list()

        for detection_path in self.detection_paths_list:
            # Open and read the detection file
            with open(detection_path, "r") as detection_file:
                detection_file_contents = detection_file.read()

            # Parse the detection and normalize fields and add to our list of detections.
            detections.append(self.__parse_and_normalize(detection_file_contents))

        return detections

    def generate_excel(self, detection_list: list) -> None:

        # Create pandas data frame from the list of dicts containing the detection info
        detection_df = pd.DataFrame(detection_list)

        # Write data frame to excel file
        detection_df.to_excel(self.output_file, index=False)

    def __parse_and_normalize(self, detection_file_contents: str) -> dict:

        # Convert detection logic to multi-line string.
        raw_detection_data = detection_file_contents.replace("detection:", "detection: |", 1)

        # Convert yaml to dict.
        raw_detection_data = yaml.load(raw_detection_data, yaml.FullLoader)

        # Normalize field names
        normalized_detection_data = dict()
        normalized_detection_data["Platform"] = "\n\n".join(raw_detection_data.get("logsource", {}).get("product", ""))
        normalized_detection_data["Severity"] = raw_detection_data.get("level", "")
        normalized_detection_data["Name"] = raw_detection_data.get("title", "")
        normalized_detection_data["Description"] = raw_detection_data.get("description", "")
        normalized_detection_data["False Positives"] = (
            "\n\n".join(raw_detection_data.get("falsepositives", ""))
            if isinstance(raw_detection_data.get("falsepositives", ""), list)
            else raw_detection_data.get("falsepositives", "")
        )
        if self.logic:
            normalized_detection_data["Detection Logic"] = raw_detection_data["detection"]

        normalized_detection_data["Reference Links"] = "\n\n".join(raw_detection_data.get("references", ""))

        return normalized_detection_data


def main():

    detection_info = DetectionData(input_dir, output_file, args.logic)
    detection_data = detection_info.fetch()
    detection_info.generate_excel(detection_data)


if __name__ == "__main__":

    parser = argparse.ArgumentParser(
        description="Export Detection Information from Sigma formatted detection Files",
        formatter_class=argparse.ArgumentDefaultsHelpFormatter,
    )

    parser.add_argument(
        "-i",
        "--input",
        type=str,
        required=True,
        help="Directory path to Active Detections",
    )

    parser.add_argument(
        "-o",
        "--output",
        type=str,
        required=True,
        help="Path to Output Excel File",
    )

    parser.add_argument(
        "-l",
        "--logic",
        action="store_true",
        default=False,
        help="Switch to toggle the full logic export",
    )

    args = parser.parse_args()

    # check if the passed input directory is valid
    if not Path.resolve(Path(args.input)).is_dir():
        parser.error(f"Path to input directory ({args.input}) of detection files is invalid.")
    else:
        input_dir = Path.resolve(Path(args.input))

    # check if the passed in output file path is to a valid directory
    if not Path.resolve(Path(args.output)).parent.is_dir():
        parser.error(f"Path to directory ({args.output}) for output file is invalid.")
    else:
        output_file = Path.resolve(Path(args.output))

    # Execute code
    main()
