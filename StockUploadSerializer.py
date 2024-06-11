from polars import DataFrame, read_excel, from_dicts

rawData = read_excel("MAPPING_TEMPLATE.xlsx",sheet_name='LX02_DATA',engine = 'openpyxl').select( ['OWNER','OWNER_ROLE','ENTITELED','ENTITLED_ROLE',
                                                                         'MATNR','HUTYP','PMAT','GR_DATE','VFDAT','LGPLA',
                                                                         'SSCC','UNIT','QUAN','CAT','EXTNO','Batch'])
rownumber = 1
serializedData = []
for rawDataRow in rawData.to_dicts():
    if rawDataRow['SSCC'] != None:
        #HU
        newDataRowHU = {'POSTYPE':"1",
                        'MATNR':"",
                        'OWNER':"",
                        'OWNER_ROLE':"",
                        'BATCH':"",
                        'CAT':"",
                        'STOCK_DOCCAT':"",
                        'STOCK_DOCNO':"",
                        'STOCK_ITMNO':"",
                        'STOCK_USAGE':"",
                        'ENTITELED':"",
                        'ENTITLED_ROLE':"",
                        'COO':"",
                        'QUAN':"",
                        'UNIT':"",
                        'HUTYP':rawDataRow['HUTYP'],
                        'LGPLA':rawDataRow['LGPLA'],
                        'GR_DATE':"",
                        'GR_TIME':"",
                        'VFDAT':"",
                        'PMAT':rawDataRow['PMAT'],
                        'EXTNO':rawDataRow['EXTNO'],
                        'HUIDENT':rawDataRow['SSCC'],
                        'PARHUIDENT':rawDataRow['SSCC'],
                        'TOPHUIDENT':rawDataRow['SSCC'],
                        'ROW':rownumber,
                        'REFROW':"",
                        'G_WEIGHT':"",'N_WEIGHT':"",'UNIT_GW':"",'T_WEIGHT':"",'UNIT_TW':"",'G_VOLUME':"",'N_VOLUME':"",'UNIT_GV':"",'T_VOLUME':"",'UNIT_TV':"",'G_CAPA':"",'N_CAPA':"",'T_CAPA':"",'LENGTH':"",'WIDTH':"",'HEIGHT':"",'UNIT_LWH':"",'MAX_WEIGHT':"",'TOLW':"",'TARE_VAR':"",'MAX_VOLUME':"",'TOLV':"",'CLOSED_PACKAGE':"",'MAX_CAPA':"",'MAX_LENGTH':"",'MAX_WIDTH':"",'MAX_HEIGHT':"",'UNIT_MAX_LWH':"",'SERNR':"",'CWQUAN':"",'CWUNIT':"",'CWEXACT':"",'LOGPOS':"",'UII':"",'AMOUNT_LC':"",'DUMMY_ISU':"",'ZEUGN':""}
        serializedData.append(newDataRowHU)
        rownumber += 1
        #Item in HU
        newDataRowHU = {'POSTYPE':"4",
                        'MATNR':rawDataRow['MATNR'],
                        'OWNER':rawDataRow['OWNER'],
                        'OWNER_ROLE':rawDataRow['OWNER_ROLE'],
                        'BATCH':rawDataRow['Batch'],
                        'CAT':rawDataRow['CAT'],
                        'STOCK_DOCCAT':"",
                        'STOCK_DOCNO':"",
                        'STOCK_ITMNO':"",
                        'STOCK_USAGE':"",
                        'ENTITELED':rawDataRow['ENTITELED'],
                        'ENTITLED_ROLE':rawDataRow['ENTITLED_ROLE'],
                        'COO':"",
                        'QUAN':rawDataRow['QUAN'],
                        'UNIT':rawDataRow['UNIT'],
                        'HUTYP':"",
                        'LGPLA':rawDataRow['LGPLA'],
                        'GR_DATE':rawDataRow['GR_DATE'],
                        'GR_TIME':"",
                        'VFDAT':rawDataRow['VFDAT'],
                        'PMAT':"",
                        'EXTNO':"",
                        'HUIDENT':rawDataRow['SSCC'],
                        'PARHUIDENT':rawDataRow['SSCC'],
                        'TOPHUIDENT':rawDataRow['SSCC'],
                        'ROW':rownumber,
                        'REFROW':newDataRowHU['ROW'],
                        'G_WEIGHT':"",'N_WEIGHT':"",'UNIT_GW':"",'T_WEIGHT':"",'UNIT_TW':"",'G_VOLUME':"",'N_VOLUME':"",'UNIT_GV':"",'T_VOLUME':"",'UNIT_TV':"",'G_CAPA':"",'N_CAPA':"",'T_CAPA':"",'LENGTH':"",'WIDTH':"",'HEIGHT':"",'UNIT_LWH':"",'MAX_WEIGHT':"",'TOLW':"",'TARE_VAR':"",'MAX_VOLUME':"",'TOLV':"",'CLOSED_PACKAGE':"",'MAX_CAPA':"",'MAX_LENGTH':"",'MAX_WIDTH':"",'MAX_HEIGHT':"",'UNIT_MAX_LWH':"",'SERNR':"",'CWQUAN':"",'CWUNIT':"",'CWEXACT':"",'LOGPOS':"",'UII':"",'AMOUNT_LC':"",'DUMMY_ISU':"",'ZEUGN':""}
        serializedData.append(newDataRowHU)
        rownumber += 1
    else:
        newDataRowHU = {'POSTYPE':"1",
                        'MATNR':rawDataRow['MATNR'],
                        'OWNER':rawDataRow['OWNER'],
                        'OWNER_ROLE':rawDataRow['OWNER_ROLE'],
                        'BATCH':rawDataRow['Batch'],
                        'CAT':rawDataRow['CAT'],
                        'STOCK_DOCCAT':"",
                        'STOCK_DOCNO':"",
                        'STOCK_ITMNO':"",
                        'STOCK_USAGE':"",
                        'ENTITELED':rawDataRow['ENTITELED'],
                        'ENTITLED_ROLE':rawDataRow['ENTITLED_ROLE'],
                        'COO':"",
                        'QUAN':rawDataRow['QUAN'],
                        'UNIT':rawDataRow['UNIT'],
                        'HUTYP':"",
                        'LGPLA':rawDataRow['LGPLA'],
                        'GR_DATE':rawDataRow['GR_DATE'],
                        'GR_TIME':"",
                        'VFDAT':rawDataRow['VFDAT'],
                        'PMAT':"",
                        'EXTNO':"",
                        'HUIDENT':rawDataRow['SSCC'],
                        'PARHUIDENT':rawDataRow['SSCC'],
                        'TOPHUIDENT':rawDataRow['SSCC'],
                        'ROW':rownumber,
                        'REFROW':"",
                        'G_WEIGHT':"",'N_WEIGHT':"",'UNIT_GW':"",'T_WEIGHT':"",'UNIT_TW':"",'G_VOLUME':"",'N_VOLUME':"",'UNIT_GV':"",'T_VOLUME':"",'UNIT_TV':"",'G_CAPA':"",'N_CAPA':"",'T_CAPA':"",'LENGTH':"",'WIDTH':"",'HEIGHT':"",'UNIT_LWH':"",'MAX_WEIGHT':"",'TOLW':"",'TARE_VAR':"",'MAX_VOLUME':"",'TOLV':"",'CLOSED_PACKAGE':"",'MAX_CAPA':"",'MAX_LENGTH':"",'MAX_WIDTH':"",'MAX_HEIGHT':"",'UNIT_MAX_LWH':"",'SERNR':"",'CWQUAN':"",'CWUNIT':"",'CWEXACT':"",'LOGPOS':"",'UII':"",'AMOUNT_LC':"",'DUMMY_ISU':"",'ZEUGN':""}
        serializedData.append(newDataRowHU)
        rownumber += 1
df = from_dicts(serializedData)
print(df)
df.write_csv('stock_upload.csv')



