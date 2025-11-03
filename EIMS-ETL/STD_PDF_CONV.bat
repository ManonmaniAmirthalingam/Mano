@ECHO OFF
CLS
ECHO You are about to execute the Non CA PDF Generation Process
"E:\Program Files\Microsoft SQL Server\130\DTS\Binn\DTExec.exe" /File "D:\ETL_PACKAGE\ETLPackages_17Oct2018\STD_PDF_CONV.dtsx" /conf "D:\ETL_PACKAGE\ETLPackages_17Oct2018\STD_PDF_CONV_CONFIG.dtsConfig" 