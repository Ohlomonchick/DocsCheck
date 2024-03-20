import os
import sys

import pathlib


def main():
    print(pathlib.Path(__file__).parent.resolve())
    args = sys.argv[1:]
    print(os.getcwd())
    print("hello-world")