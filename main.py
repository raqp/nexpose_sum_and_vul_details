import os
import yaml
import platform
from sys import argv
from handlers.nexpose_sum_and_vul_detail import ExecutiveSummary


def config_loader():
    config_file = os.path.join(os.path.dirname(os.path.abspath(__file__)), "config.yaml")
    with open(config_file) as yaml_file:
        conf = yaml.load(yaml_file, Loader=yaml.FullLoader)
    return conf


def get_run_params():
    try:
        return argv[1], argv[2], argv[3]
    except IndexError:
        print("Incorrect run params.")
        exit(127)


def get_platform():
    return platform.system().lower()


def start():
    config = config_loader()
    source_file, destination_path, input_document_type = get_run_params()
    os_platform = get_platform()
    summary = ExecutiveSummary(config=config, source_file=source_file, destination_path=destination_path,
                               input_document_type=input_document_type, os_platform=os_platform)
    summary.start()


if __name__ == '__main__':
    start()

