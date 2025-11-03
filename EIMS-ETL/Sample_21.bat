@ECHO OFF
CLS
ECHO You are about to execute the CA HIP Daily Process
"E:\Program Files\Microsoft SQL Server\130\DTS\Binn\DTExec.exe" /File "Z:\EIMS Changes\EBT_CA_PDF_CONV_P21_HIP.dtsx" /conf "Z:\EIMS Changes\EBT_CA_PDF_CONV_P21_HIP_Config.dtsConfig"


pause

 
