from excelsheet2html import parsing_formula


if __name__ == "__main__":
    f = '=2*(A4-7*SUM(B1:B4))-C5'
    print(parsing_formula(f))
