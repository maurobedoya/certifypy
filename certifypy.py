"""
Version: 1.0
Generator of poster presentation certificates, attendance at conferences, 
organizing committee and juror participation and awards.
This script reads an input file (input.dat) with all the options and will 
generate the requested certificate.
An excel file is also required with the list of attendees and additional 
information such as the institution, name of the work or talk or award 
obtained.
"If this script is useful for you, just be happy, it's free! :)"
Contact: 
Mauricio Bedoya
maurobedoyat@gmail.com"
"""
from __future__ import print_function

import argparse
import os
from posix import sched_param
import subprocess
import sys
import glob
from tabnanny import check
from typing_extensions import TypeAlias
import numpy as np
from os import path, PathLike, supports_fd, write
import pandas as pd

from dataclasses import dataclass
from typing import Dict, List, Optional, Set, TextIO, Tuple
import configparser
import random

# Font and openpyxl object
# import cv2 as cv
# import openpyxl

from PIL import Image, ImageDraw, ImageFont
import textwrap


def parse_args(argv):
    conf_parser = argparse.ArgumentParser(
        description=__doc__,
        formatter_class=argparse.RawDescriptionHelpFormatter,
        add_help=False,
    )
    conf_parser.add_argument(
        "-i", "--input", help="Specify a configuration file", metavar="FILE"
    )
    args, remaining_argv = conf_parser.parse_known_args()

    # defaults = {"participants_data": "$SCHRODINGER"}
    defaults = {}

    print(
        """
    █▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀█
    █      ┌─┐┌─┐┬─┐┌┬┐┬┌─┐┬ ┬┌─┐┬ ┬      █
    █      │  ├┤ ├┬┘ │ │├┤ └┬┘├─┘└┬┘      █
    █      └─┘└─┘┴└─ ┴ ┴└   ┴ ┴   ┴  v1.0 █
    █▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄█

    Generator of poster presentation certificates, attendance at conferences, 
    organizing committee and juror participation and awards.
    1. Pass an input file to the script.
    2. Execute as: python certifypy.py -i input.dat
    3. The full list of configuration file options can be found at: 
    https://github.com/maurobedoya/certifypy

    "If this script is useful for you, just be happy, it's free! :)"
    Contact: 
    Mauricio Bedoya
    maurobedoyat@gmail.com"
    
    """
    )

    try:
        if args.input:
            config = configparser.ConfigParser()
            config.read([args.input])
            defaults.update(dict(config.items("settings")))
            # defaults.update(dict(config.items("layout")))
    except:
        raise ValueError("You must specify an input file")

    # Parse rest of arguments
    # Don't suppress add_help here so it will handle -h
    parser = argparse.ArgumentParser(
        # Inherit options from config_parser
        parents=[conf_parser]
    )
    parser.set_defaults(**defaults)
    args = parser.parse_args(remaining_argv)
    layout_opts = {}
    layout_opts.update(dict(config.items("layout")))
    template_name = vars(args)["template"].split("/")[-1]
    template = vars(args)["template"]
    participants_data = vars(args)["participants_data"]
    template_path = os.path.abspath(vars(args)["template"])
    fonts_folder = os.path.abspath(vars(args)["fonts_folder"])
    info_opts = {}
    info_opts.update(dict(config.items("info")))

    return (
        Args(**vars(args)),
        LayoutOptions(
            layout_opts,
            template,
            template_name,
            participants_data,
            template_path,
            fonts_folder,
        ),
        template_path,
        InfoOptions(info_opts),
    )


@dataclass
class Args:
    input: str
    basename: str
    participants_data: str
    workdir: str = ""
    fonts_folder: str = ""
    template: str = ""


@dataclass
class LayoutOptions:
    paper_size: Optional[str] = "A4"
    orientation: Optional[str] = "vertical"
    basename: Optional[str] = "None"
    custom_size: Optional[str] = "10.0 10.0"

    def __init__(
        self,
        opts: Dict[str, str],
        template: str,
        template_name: str,
        participants_data: str,
        template_path: str,
        fonts_folder: str,
    ) -> None:
        self.opts = opts
        self.template = template
        self.template_name = template_name
        self.basename = str(template_name).split(".")[0]
        self.participants_data = participants_data
        self.template_path = template_path
        self.fonts_folder = fonts_folder
        for key in self.opts:
            setattr(self, key, self.opts[key])

    def __getattr__(self, item):
        if item not in self.opts:
            return None
        return self.opts[item]


@dataclass
class InfoOptions:
    title: Optional[str] = "None"
    title_coords: Optional[str] = "None"
    title_font: Optional[str] = "None"
    title_font_size: Optional[str] = "None"
    title_font_color: Optional[str] = "None"
    title_font_bold: Optional[str] = "None"
    title_font_italic: Optional[str] = "None"
    title_font_underline: Optional[str] = "None"
    title_font_shadow: Optional[str] = "None"
    title_font_outline: Optional[str] = "None"

    subtitle: Optional[str] = "None"
    subtitle_coords: Optional[str] = "None"
    subtitle_font: Optional[str] = "None"
    subtitle_font_size: Optional[str] = "None"
    subtitle_font_color: Optional[str] = "None"
    subtitle_font_bold: Optional[str] = "None"
    subtitle_font_italic: Optional[str] = "None"
    subtitle_font_underline: Optional[str] = "None"
    subtitle_font_shadow: Optional[str] = "None"
    subtitle_font_outline: Optional[str] = "None"

    participant_name: Optional[bool] = True
    participant_name_coords: Tuple[float, float] = (0, 0)
    participant_name_font: Optional[str] = "None"
    participant_name_font_size: Optional[str] = "None"
    participant_name_font_color: Optional[str] = "None"

    participant_affiliation: Optional[bool] = True
    participant_affiliation_coords: Tuple[float, float] = (0, 0)
    participant_affiliation_font: Optional[str] = "None"
    participant_affiliation_font_size: Optional[str] = "None"
    participant_affiliation_font_color: Optional[str] = "None"

    participant_work_title: Optional[bool] = True
    participant_work_title_coords: Tuple[float, float] = (0, 0)
    participant_work_title_font: Optional[str] = "None"
    participant_work_title_font_size: Optional[str] = "None"
    participant_work_title_font_color: Optional[str] = "None"

    date: Optional[bool] = True
    date_coords: Tuple[float, float] = (0, 0)
    date_font: Optional[str] = "None"
    date_font_size: Optional[str] = "None"
    date_font_color: Optional[str] = "None"

    attendant_title: Optional[str] = "None"
    attendant_title_coords: Tuple[float, float] = (0, 0)
    attendant_title_font: Optional[str] = "None"
    attendant_title_font_size: Optional[str] = "None"
    attendant_title_font_color: Optional[str] = "None"

    attendant_text: Optional[str] = "Attendant"
    attendant_text_coords: Tuple[float, float] = (0, 0)
    attendant_text_font: Optional[str] = "None"
    attendant_text_font_size: Optional[str] = "None"
    attendant_text_font_color: Optional[str] = "None"

    poster_title: Optional[str] = "None"
    poster_title_coords: Tuple[float, float] = (0, 0)
    poster_title_font: Optional[str] = "None"
    poster_title_font_size: Optional[str] = "None"
    poster_title_font_color: Optional[str] = "None"

    poster_text: Optional[str] = "poster"
    poster_text_coords: Tuple[float, float] = (0, 0)
    poster_text_font: Optional[str] = "None"
    poster_text_font_size: Optional[str] = "None"
    poster_text_font_color: Optional[str] = "None"

    # Run infos
    run_preparation: Optional[str] = "false"
    run_infos: Optional[str] = "false"

    def __init__(self, opts: Dict) -> None:
        self.opts = opts
        for key in self.opts:
            setattr(self, key, self.opts[key])

    def __getattr__(self, item):
        if item not in self.opts:
            return None
        return self.opts[item]


def check_output_folder(folder_name: str):
    if path.isdir(folder_name):
        raise ValueError(f"Folder '{folder_name}' exists, remove it before to continue")
    else:
        os.makedirs(folder_name)
        os.chdir(folder_name)


def draw_multiple_line_text(image, text, font, text_color, text_start_height):
    """
    From unutbu on [python PIL draw multiline text on image](https://stackoverflow.com/a/7698300/395857)
    """
    draw = ImageDraw.Draw(image)
    image_width, image_height = image.size
    y_text = text_start_height
    lines = textwrap.wrap(text, width=40)
    for line in lines:
        line_width, line_height = font.getsize(line)
        draw.text(
            ((image_width - line_width) / 2, y_text), line, font=font, fill=text_color
        )
        y_text += line_height


def certificate(
    settings: Args,
    info: InfoOptions,
    template,
    name: str,
    fonts,
    affiliation: str,
    type_cert: str,
    work_title: str,
):
    image = Image.open(template)
    width, height = image.size
    draw = ImageDraw.Draw(image)
    if type_cert == "attendant":
        text = info.attendant_text
        font_title = os.path.join(fonts, info.attendant_title_font)
        font_body = os.path.join(fonts, info.attendant_text_font)
        font_t = ImageFont.truetype(font_title, int(info.attendant_title_font_size))
        font_b = ImageFont.truetype(font_body, int(info.attendant_text_font_size))
        coords_title = (
            int(float(info.attendant_title_coords.split(",")[0]) * width),
            int(float(info.attendant_title_coords.split(",")[1]) * height),
        )
        coords_body = (
            int(float(info.attendant_text_coords.split(",")[0]) * width),
            int(float(info.attendant_text_coords.split(",")[1]) * height),
        )
        # Title
        draw.text(
            coords_title,
            info.attendant_title,
            fill=info.attendant_title_font_color,
            anchor="ms",
            font=font_t,
            align="center",
        )
        # Body
        draw_multiple_line_text(
            image=image,
            text=text,
            font=font_b,
            text_color=info.attendant_text_font_color,
            text_start_height=coords_body[1],
        )
    if type_cert == "poster":
        work_text = f'"{work_title}"'
        text = info.poster_text
        font_title = os.path.join(fonts, info.poster_title_font)
        font_body = os.path.join(fonts, info.poster_text_font)
        font_work = os.path.join(fonts, info.participant_work_title_font)

        font_t = ImageFont.truetype(font_title, int(info.poster_title_font_size))
        font_b = ImageFont.truetype(font_body, int(info.poster_text_font_size))
        font_w = ImageFont.truetype(
            font_work, int(info.participant_work_title_font_size)
        )
        coords_title = (
            int(float(info.poster_title_coords.split(",")[0]) * width),
            int(float(info.poster_title_coords.split(",")[1]) * height),
        )
        coords_body = (
            int(float(info.poster_text_coords.split(",")[0]) * width),
            int(float(info.poster_text_coords.split(",")[1]) * height),
        )
        coords_work = (
            int(float(info.participant_work_title_coords.split(",")[0]) * width),
            int(float(info.participant_work_title_coords.split(",")[1]) * height),
        )
        # Title
        draw.text(
            coords_title,
            info.poster_title,
            fill=info.poster_title_font_color,
            anchor="ms",
            font=font_t,
            align="center",
        )
        # Body
        draw_multiple_line_text(
            image=image,
            text=text,
            font=font_b,
            text_color=info.poster_text_font_color,
            text_start_height=coords_body[1],
        )
        # Work text wit quotes
        draw_multiple_line_text(
            image=image,
            text=work_text,
            font=font_w,
            text_color=info.participant_work_title_font_color,
            text_start_height=coords_work[1],
        )

    font_subtitle = os.path.join(fonts, info.subtitle_font)
    font_name = os.path.join(fonts, info.participant_name_font)
    font_affil = os.path.join(fonts, info.participant_affiliation_font)
    font_s = ImageFont.truetype(font_subtitle, int(info.subtitle_font_size))
    font_n = ImageFont.truetype(font_name, int(info.participant_name_font_size))
    font_a = ImageFont.truetype(font_affil, int(info.participant_affiliation_font_size))

    coords_subtitle = (
        int(float(info.subtitle_coords.split(",")[0]) * width),
        int(float(info.subtitle_coords.split(",")[1]) * height),
    )

    coords_name = (
        int(float(info.participant_name_coords.split(",")[0]) * width),
        int(float(info.participant_name_coords.split(",")[1]) * height),
    )

    coords_affil = (
        int(float(info.participant_affiliation_coords.split(",")[0]) * width),
        int(float(info.participant_affiliation_coords.split(",")[1]) * height),
    )

    # Name
    draw.text(
        coords_name,
        name,
        fill=info.participant_name_font_color,
        anchor="ms",
        font=font_n,
        align="center",
    )
    # Affiliation in multiple lines
    draw_multiple_line_text(
        image=image,
        text=affiliation,
        font=font_a,
        text_color=info.participant_affiliation_font_color,
        text_start_height=coords_affil[1],
    )
    # Subtitle
    draw.text(
        coords_subtitle,
        info.subtitle,
        fill=info.subtitle_font_color,
        anchor="ms",
        font=font_s,
        align="center",
    )

    image.save(f"{settings.basename}_{name.replace(' ', '_')}_{type_cert}.png")


def main(argv):
    settings, layout_opts, template_path, info = parse_args(argv)
    fonts = layout_opts.fonts_folder
    template = template_path
    # print(settings)
    # print(layout_opts)
    # print(info)
    # print(layout_opts.fonts_folder)

    participants_data = pd.read_excel(settings.participants_data, header=0)
    # print(participants_data)
    participants_data.fillna(0, inplace=True)
    # print(participants_data)
    check_output_folder(settings.workdir)
    for index, row in participants_data.iterrows():
        name = row["NAME"]
        affiliation = row["AFFILIATION"]
        type_cert = "attendant"
        print(f"Processing certificate for {name} ...")
        certificate(
            settings=settings,
            info=info,
            template=template,
            name=name,
            fonts=fonts,
            affiliation=affiliation,
            type_cert=type_cert,
            work_title="",
        )
        if row["POSTER"] != 0:
            poster_title = row["POSTER"]
            type_cert = "poster"
            certificate(
                settings=settings,
                info=info,
                template=template,
                name=name,
                fonts=fonts,
                affiliation=affiliation,
                type_cert=type_cert,
                work_title=poster_title,
            )

        if row["TALK"] != "":
            talk_title = row["TALK"]
        else:
            talk_title = "None"

        if row["AWARD"] != "":
            award_title = row["AWARD"]
        else:
            award_title = "None"
        if row["ROLE"] != "":
            attendant_role = row["ROLE"]


if __name__ == "__main__":
    main(sys.argv[1:])