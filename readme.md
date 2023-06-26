# WMI-Parser

## TLDR

This is an updated version of woanware's WMI-Parser, which can be found [here](https://github.com/woanware/wmi-parser). That tool was a C# rewrite of davidpany's WMI_Forensics Python script, found [here](https://github.com/davidpany/WMI_Forensics)

## Usage

Below is an updated version of the original example from the previous WMI-Parser repo:
```
.\wmi-parser.exe -i C:\temp\testdata\OBJECTS.DATA -o c:\temp\testdata

wmi-parser v1.0.0

Author: Mark Woan / woanware (markwoan@gmail.com)
https://github.com/woanware/wmi-parser

Updated By: Andrew Rathbun / https://github.com/AndrewRathbun
https://github.com/AndrewRathbun/wmi-parser

  SCM Event Log Consumer-SCM Event Log Filter - (Common binding based on consumer and filter names, possibly legitimate)
    Consumer: NTEventLogEventConsumer ~ SCM Event Log Consumer ~ sid ~ Service Control Manager

    Filter:
      Filter Name : SCM Event Log Filter
      Filter Query: select * from MSFT_SCMEventLogEvent

  BadStuff-DeviceDocked

    Name: BadStuff
    Type: CommandLineEventConsumer
    Arguments: powershell.exe -NoP Start-Process ('badstuff.exe')

    Filter:
      Filter Name : DeviceDocked
      Filter Query: SELECT * FROM Win32_SystemConfigurationChangeEvent 
```

Below is an example of parsing an infected `OBJECTS.DATA` file, the sample of which can be found [here](https://github.com/AndrewRathbun/DFIRArtifactMuseum/tree/main/Windows/WBEM/Win10/APTSimulatorVM_Defender):

![image](https://github.com/AndrewRathbun/WMI-Parser/assets/36825567/0cd454cf-3855-4e58-b2fe-75a6cdbe2980)

## Output

This will output to TSV, which can ingested into a tool like Timeline Explorer for easy analysis.

