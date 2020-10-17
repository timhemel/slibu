# Slide Builder

Build PowerPoint slides from Markdown.

## Introduction

Slibu (short for slide builder) lets you write slides in Markdown and generates a
PowerPoint file from it. You can now combine the inconvenience of a non-WYSIWYG editor
with the limitations of PowerPoint! Why would you want that?

* Store slides in version control.
* Let your program create slides by generating Markdown.
* Take advantage of Markdown features such as syntax highlighting.
* ...

## Usage

```
Usage: slibu [OPTIONS]

Options:
  -o, --out-file TEXT  write output to this file (default out.pptx)
  -t, --template TEXT  pptx file to use as slide template (default
                       reference.pptx)
  --help               Show this message and exit.
```

Slibu reads markdown from the standard input and expects a certain format. It
also expects a PowerPoint template (default is `reference.pptx` in the current
directory). It then writes the PowerPoint to the output file (default `out.pptx`
in the current directory).


## Markdown syntax

Slibu uses [mistune](https://github.com/lepture/mistune) to parse Markdown and makes
uses of its *directives* feature. Furthermore, it expects a certain structure:

```
# This starts a presentation part

## This starts a slide

Content goes here.
```


