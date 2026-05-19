import re

def update_formula(formula, source_row, target_row):
    if not isinstance(formula, str) or not formula.startswith('='):
        return formula
    
    delta = target_row - source_row
    
    def repl(m):
        prefix_col = m.group(1)
        col = m.group(2)
        prefix_row = m.group(3)
        row = int(m.group(4))
        
        # If the row is absolute (e.g. $A$5), don't change it
        if prefix_row:
            return m.group(0)
            
        # Only update if the row matches the expected source row or if we just want to shift all relative rows
        # Actually, in Excel, if you copy a cell down 1 row, ALL relative row references increase by 1.
        return f"{prefix_col}{col}{prefix_row}{row + delta}"
        
    return re.sub(r'(\$?)([A-Z]+)(\$?)(\d+)', repl, formula)

formula = '=IF(COUNTIFS($A$2:A5,Pedagogico_1[[#This Row],[REGIONAL]],$G$2:G5,Pedagogico_1[[#This Row],[ZONA 2026.]])=1,SUMIFS(Pedagogico_1[VALOR TOTAL INICIAL APORTE ICBF POR SERVICIO\n$],Pedagogico_1[REGIONAL],Pedagogico_1[[#This Row],[REGIONAL]],Pedagogico_1[ZONA 2026.],Pedagogico_1[[#This Row],[ZONA 2026.]]),0)'
print(update_formula(formula, 5, 6))
