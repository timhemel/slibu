#!/usr/bin/env python3

import sys
import mistune
from mistune import AstRenderer
from ImageDirective import ImageDirective
from BoxDirective import BoxDirective
from TitleDirective import TitleDirective
from SlideDirective import SlideDirective
from TextStyleDirective import TextStyleDirective
from SetStyleDirective import SetStyleDirective
from mistune.directives.admonition import Admonition
from pptx_writer import PPTXWriter

markdown = mistune.create_markdown(renderer=mistune.AstRenderer(), plugins=['table', ImageDirective(), BoxDirective(), TitleDirective(), TextStyleDirective(), SetStyleDirective(), SlideDirective()])
# markdown = mistune.create_markdown(renderer=mistune.AstRenderer(), plugins=['table', Admonition()])
r = PPTXWriter()
r.feed(markdown(sys.stdin.read()))
r.save("test.pptx")

