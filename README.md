# Build a block schedule
## About

This Python 3 programme reads a Microsoft Project XML file and creates a Scalable Vector Graphics file where all project tasks are represented by __blocks__.
Summary tasks containing tasks are represented by nested blocks, where the critical path is displayed by red blocks.

The number of levels to be converted to `SVG` can be entered.

## Running

The routine is called using:

`CreateBlockSchedule(<FileName> [, <List of levels>]):`

