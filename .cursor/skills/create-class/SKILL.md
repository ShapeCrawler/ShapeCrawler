---
name: class
description: Use this skill when creating a new class or asking how to model a new class in ShapeCrawler.
---

# Create Class

## Rule (Key Design Patterns)

When creating a new class:

- The class must represent a real-world entity or concept (noun).
- The constructor must encapsulate the entity’s coordinates: store the properties or underlying object the class uses to refer to that entity.
- The class must encapsulate all logic required to produce its result; do not rely on pre-calculated data from callers if it can be derived from dependencies.

## Quick Checks

- Class name is a noun, not “-er/-or/-service/-manager”.
- Constructor captures the underlying object/properties that identify the entity.
- Callers pass minimal inputs; the class derives the rest from its dependencies.
