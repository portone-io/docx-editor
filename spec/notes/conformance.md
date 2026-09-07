# Conformance classes and namespace prefixes

## A conformance class is a property of the package, not of a part

ECMA-376 defines two document conformance classes.
Office Open XML Strict is Part 1 §2.1; Office Open XML Transitional is Part 4 §2.1.
Both describe the same vocabulary, and a package belongs wholly to one of them: the class decides the namespace every part's root carries and the relationship type the package reaches each part through.

For the main document part, Part 1 §11.3.10 gives the Strict root namespace `http://purl.oclc.org/ooxml/wordprocessingml/main` and the source relationship `http://purl.oclc.org/ooxml/officeDocument/relationships/officeDocument`.
Part 4 §9.2.10 gives the Transitional pair, `http://schemas.openxmlformats.org/wordprocessingml/2006/main` and `http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument`.
A package therefore says which class it belongs to twice, and a reader can tell from either.

## Refusing Strict is conforming

Application conformance is stated per class in the same words in both parts, Part 1 §2.2 and Part 4 §2.2: a conforming consumer shall not reject any conforming document of at least one document conformance class.
Reading and writing Transitional alone is a conforming consumer and producer of WML Transitional.
This editor therefore refuses a Strict package outright, under its own import error code, rather than reading part of it and writing something back that belongs to neither class.
Word saves Transitional by default, so a Strict package is rare in practice, and the refusal is a named one rather than a later failure over markup that did not resolve.

## A namespace prefix is the producer's choice

Prefixes carry no meaning of their own; the namespace a prefix is bound to does.
Part 1 Annex D is explicit that the prefixes used throughout the standard's own examples are an illustration convention, listed there precisely because the examples leave the bindings out.
Nothing in the standard requires a WordprocessingML element to be written `w:`, and a document may bind the namespace to another prefix or as its default.

Reading follows from that: an element or attribute that arrived in a file is looked up by its local part and its namespace, never by the spelling a producer chose.

Writing cannot: a prefix has to be spelled out, and this editor spells `w`.
Markup it writes is spliced into a part the document brought, so the binding has to be in scope where the markup lands.
Rather than declaring the prefix on every element written, the root of each part written binds it, and a main part whose root does not bind `w` to the Transitional namespace is refused when the file is opened.
That refusal is about what can be written back, not about what the standard allows, so it carries the code for markup this editor cannot write rather than the conformance one.
