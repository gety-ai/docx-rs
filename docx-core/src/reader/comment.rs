// Comment deserialization is handled by serde Deserialize impl on Comment struct.
// The ElementReader for Comment is no longer needed - Comments uses quick-xml serde
// to deserialize its children directly.
