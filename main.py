import os
import argparse

import pptx
from pptx.presentation import Presentation

import slide_index
import slide_num


def get_args():
    parser = argparse.ArgumentParser(
        prog="ProgramName",
        description='What the program does',
        epilog='Text at the bottom of help'
    )
    parser.add_argument('filename')
    args = parser.parse_args()
    return args


def main(filename: str):
    new_name = os.path.splitext(filename)[0] + '.pi.pptx'
    prs: Presentation = pptx.Presentation(filename)

    index = slide_index.add_index(prs)
    slide_index.add_index_sidebar(prs, index)
    slide_num.add_total_slide_num(prs)
    prs.save(new_name)


if __name__ == '__main__':
    args = get_args()
    filename = args.filename
    main(filename)
