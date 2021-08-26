# excelsheet2html

Export table from excel sheet as dynamic HTML page. The page looks like a part of sheet and contains dynamic functions such as: resetting columns, previewing formulas, modifying cell values.

You can preview output of first example on: [https://codepen.io/lwysocki/pen/OJmBvEM](https://codepen.io/lwysocki/pen/OJmBvEM)

# Install

Not yet ready

# Usage

Check examples

Save as html

```python
from excelsheet2html import excelsheet2html

output = excelsheet2html("Book1.xlsx")
f = open("Book2.html", 'w')
f.write(output.html)
f.close()
```

or parse formula

```python
from excelsheet2html import parsing_formula

f = '=2*(A4-7*SUM(B1:B4))-C5'
print(parsing_formula(f))
```
