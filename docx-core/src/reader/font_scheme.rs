// FontScheme deserialization is handled by serde Deserialize impl on FontScheme struct.
// The ElementReader for FontScheme is no longer needed - Theme uses quick-xml serde
// to deserialize FontScheme directly as a nested child.
