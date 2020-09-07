
# README


*.properties files point to their respective EXCEL_SHEET with a relative path, separated by \\. 

------------------
File Structure
------------------

- EUSES\
  - configuration_files
    - cs101
      - *.properties --> ../../spreadsheets/cs101/SEEDED/*FAULT*.xls(x)
      
    - database
    - ...
    
  - spreadsheets
    - cs101
      - original
        - *.xls(x)
      - SEEDED
        - *FAULT*.xls(x)
      
    - database
    - ...
