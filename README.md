# Exify Office
A simple Bash wrapper for EXIficient and EXIProcessor to compress Office Open XML documents using Efficient XML Interchange.

## Quickstart Guide

__With Subversion__

1. Checkout local copy:

```
$ svn checkout https://github.com/hillbw/exify-office
```

2. Change into project directory:

```
$ cd exify-office
```

3. Compress the included sample file:

This will create a file called hello_world.doce ('e' for Efficient) in the same directory as the given docx file.

```
$ ./exify_office.sh samples/hello_world.docx
```

4. Decompress

This creates a file called hello_unexify.docx, again in the samples directory.

```
$ ./unexify_office.sh samples/hello_world.doce
```

## Notes

1. The scripts use the following conversions for filename extensions:

| Original Extension | Exified Extension |
|--------------------|-------------------|
| `.docx`            | `.doce`           |
| `.pptx`            | `.ppte`           |
| `.xlsx`            | `.xlse`           |

2. Even for the trivial Hello World .docx file, the roundtrip conversion may decrease the filesize. Generally, this is because when EXIficient writes EXI out to XML, it will condense empty elements to singleton notations. Office appears to use the expanded open-tag/close-tag notation in all cases.

```
<!-- Original .docx format -->
<dc:description></dc:description>

<!-- After roundtrip conversion -->
<dc:description/>
```