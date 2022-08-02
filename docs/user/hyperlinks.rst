Working with Hyperlinks
=======================

Hyperlinks are inline text objects. They always exist within a paragraph, and contain 
one or more runs. 

There are two types of hyperlinks in a word document: internal (used for bookmarks and 
internal documents references) and external (for example pointing to a website). 
Internal hyperlinks are currently not implemented - since the relevant other internal
features (such as bookmarks) are not implemented. External hyperlinks are implements.

The word document standard stores the target (URL) of hyperlinks seperately from the
main document. You'll have to use the main document object to lookup or set these 
targets. Therefore, methods for creating and updating hyperlinks require that you pass
in the main document object (see below). This API is obviously clunky and non-optimal,
but allowed adding hyperlink features without drastically changing the underlying 
library architecture.

For example, to create a hyperlink:

    >>> doc = Document()
    >>> p = doc.add_paragraph()
    >>> hyperlink = p.add_hyperlink("Click here", hyperlink_url="https://www.example.com", document=doc)

To change a hyperlink target:

    >>> hyperlink.update_external_target("https://www.foobar.com", doc)