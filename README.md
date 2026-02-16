# TOPSIS-Python Package

**Author:** Neeraj  

**Roll Number:** 102317014

---

## What is TOPSIS?

**T**echnique for **O**rder of **P**reference by **S**imilarity to **I**deal **S**olution (TOPSIS) is a multi-criteria decision analysis method. It is based on the concept that the chosen alternative should have the shortest geometric distance from the positive ideal solution (PIS) and the longest geometric distance from the negative ideal solution (NIS).

---

## Installation

You can install the package using pip:

pip install Topsis-Neeraj-102317014

---

## Example Usage

topsis data.xlsx "1,1,1,2" "+,+,-,+" result.xlsx

---

## Input File Format

The input file must:

Be a .xlsx file.

Contain at least 3 columns.

The first column should contain the object/alternative names (e.g.,M1,M2,M3).

From the 2nd column onwards, all values must be numeric.

---

## Output

The output will be a excel file containing all the original data, plus two additional columns:

Topsis Score: The calculated preference score.

Rank: The final ranking of the alternatives.

---
