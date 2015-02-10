# Exify Office
A simple Bash wrapper for EXIficient and EXIProcessor to compress Office Open XML documents using Efficient XML Interchange.

## Quickstart Guide

__With Subversion__

Checkout local copy:

```
$ svn checkout https://github.com/hillbw/exify-office
```

Change into project directory:

```
$ cd exify-office
```

Compress the included sample file:

This will create a file called schedule.xlse ('e' for Efficient) in the same directory as the given xls file.

```
$ ./exify_office.sh samples/schedule.xlsx
```

Decompress

This creates a file called hello_unexify.docx, again in the samples directory.

```
$ ./unexify_office.sh samples/schedule.xlse
```

## Notes

The scripts use the following conversions for filename extensions:

| Original Extension | Exified Extension |
|--------------------|-------------------|
| `.docx`            | `.doce`           |
| `.pptx`            | `.ppte`           |
| `.xlsx`            | `.xlse`           |

Even for the trivial `hello.docx` file, the roundtrip conversion may decrease the filesize. Generally, this is because when EXIficient writes EXI out to XML, it will condense empty elements to singleton notations. Office appears to use the expanded open-tag/close-tag notation in all cases.

```xml
<!-- Original .docx format -->
<dc:description></dc:description>

<!-- After roundtrip conversion -->
<dc:description/>
```