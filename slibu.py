#!/usr/bin/env python3

# SliBu, the slide builder

import sys
import click
import mistune
from mistune import AstRenderer
from ImageDirective import ImageDirective
from BoxDirective import BoxDirective
from TitleDirective import TitleDirective
from SlideDirective import SlideDirective
from TextStyleDirective import TextStyleDirective
from SetStyleDirective import SetStyleDirective
from pptx_writer import PPTXWriter

@click.command()
@click.option('-o', '--out-file', default='out.pptx', help='write output to this file (default out.pptx)')
@click.option('-t', '--template', default='reference.pptx', help='pptx file to use as slide template (default reference.pptx)')
def build(out_file, template):

    markdown = mistune.create_markdown(renderer=mistune.AstRenderer(),
            plugins=['table', ImageDirective(), BoxDirective(),
            TitleDirective(), TextStyleDirective(), SetStyleDirective(),
            SlideDirective()])

    r = PPTXWriter(template)
    r.feed(markdown(sys.stdin.read()))
    r.save(out_file)


