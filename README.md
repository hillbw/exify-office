# Exify Office
A simple Bash wrapper for [EXIficient](http://sourceforge.net/projects/exificient/) and [EXIProcessor](http://sourceforge.net/projects/exiprocessor/) to compress Office Open XML documents using [Efficient XML Interchange](http://www.w3.org/TR/2014/REC-exi-20140211/).

## Quickstart Guide

### Checkout

```
$ svn checkout https://github.com/hillbw/exify-office
```

### Change into project directory

```
$ cd exify-office
```

### Compress the included sample file:

This will create a file called schedule.xlse ('e' for Efficient) in the same directory as the given xls file.

```
$ ./exify_office.sh samples/schedule.xlsx
```

### Decompress

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

After a roundtrip conversion, the resulting Office file may be smaller than the original. Generally, this is because when EXIficient writes EXI out to XML, it will condense empty elements to singleton notations. Office appears to use the expanded open-tag/close-tag notation in all cases.

```xml
<!-- Original .docx format -->
<dc:description></dc:description>

<!-- After roundtrip conversion -->
<dc:description/>
```