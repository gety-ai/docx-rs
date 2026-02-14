// FontGroup deserialization is handled by serde Deserialize impl on FontGroup struct.
// The ElementReader for FontGroup is no longer needed - FontScheme uses quick-xml serde
// to deserialize FontGroup directly as nested children.
