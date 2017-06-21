# customPdfMerger
pdf merger that merges specific pages in a specific order

user input format ex: 11111111 
-8 digit number, no spaces

file name conventions:
  plain: AB12345678
  seperate page 3: AB12345678PG03
  separate first page: AB12345678PG00 or AB12345678PG01
  separate last page: AB12345678PG99
  merged one: AB12345678_UPDATED
  merged all: WB12345678

takes files from two different predefined directories, 
merges as requested and places the results in two different predefined directories
