# SAP_ECC_BDS_FORM_MS_WORD
SAP BDS Form - MS Word Tutorial

[Tutorial] Create Program Form BDS (Business Document Services) with Layout Microsoft Word
by: John Eswin Nizar

Video : https://youtu.be/UMddaS8b4OE

1. Edit Document (field) in MS Word

2. Create class name, class type, and object key in Business Document Navigator 
   ( Tcode: SBDSV1 )
   Mandatory -> Field name in MS Word = Field name in SAP (sequence)

TYPES:  BEGIN OF ty_data,
          ponumber      type ekko-ebeln,    " PO_NUMBER
          hari          type langt,         " SPELL_DAYS_PRINT_DATE
          tanggal       type spell-word,    " SPELL_DATE_PRINT_DATE
          bulan         type fcltx,         " SPELL_MONTH_PRINT_DATE
          tahun         type spell-word,    " SPELL_YEAR_PRINT_DATE
          tgl           type char10,        " PRINT_DATE
          plantdesc     type t001w-name1,   " PLANT_DESC
          vendordesc    type lfa1-name1,    " VENDOR_DESC
          vendoraddress type char255,       " VENDOR_ADDRESS
          podocdate     type C length 10,   " PO_DOC_DATE
          totamount     type char35,        " TOTAL_AMOUNT
          totppn        type char35,        " PPN 10%
          total         type char35,        " TOTAL (TOTAL_AMOUNT + PPN 10%)
          terbilang     type spell-word,    " TERBILANG
        END OF ty_data.

    m_namecol:  'PONUMBER' '1',
                'HARI' '2',
                'TANGGAL' '3',
                'BULAN' '4',
                'TAHUN' '5',
                'TGL' '6',
                'PLANTDESC' '7',
                'VENDORDESC' '8',
                'VENDORADDRESS' '9',
                'PODOCDATE' '10',
                'TOTAMOUNT' '11',
                'TOTPPN' '12',
                'TOTAL' '13',
                'TERBILANG' '14'.

3. Upload File .docx to SAP ( Tcode: OAOR )
   You must closed the file in MS Word and then upload.

4. Create ABAP Program 

5. Test
