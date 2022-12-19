# Siemens Insight

Excel macro to combine multiple reports into a single excel file. This was tested on Siemens Apogee Insight.

## Point Definition Report

|Point System Name:|Example|||
|-|-|-|-|
|Point Name:|Example|
|Point Type:|LAO|
|Supervised:|No|
|Revision Number:|2|
|Classification:|Building Automation|
|Descriptor:|VIRT Example|
|Panel Name:|PNL1| (PNL1)|
|Point Address:| -Virtual-|
|Actuator Type|L-Type|
|Slope:|1.0|Intercept:|0.0|
|COV Limit:|1.0|
|Engineering Units:|NULL|
|Analog Representation:|Float|
|# of decimal places:|2|
|Initial Value:|0.0|
|Totalization:|None|Initial Priority:|NONE|
|Enabled for RENO:|No|
|Alarm Issue Management:|No|
|Graphic Name:|<Undefined>|
|Informational Text:|<< No Text Defined >>|
|Alarm Type:|Not Alarmable|

## RENO Definition Report

|Example|NORMAL|Example Alarming Group||
|-|-|-|-|
||PRI1|Example Alarming Group||
||PRI2|Example Alarming Group||
||PRI3|Example Alarming Group||
||PRI5|Example Alarming Group||
||PRI6|Example Alarming Group||
||FAILED|Example Alarming Group||

## Trend Definition Report

|Point System Name:|Example|
|-|-|
|Supervised:|Yes|
|Revision Number:|24|
|Descriptor:|VIRT Example|
|Definition 1|
|Trend Every:|15 minutes|
|Samples at Panel:|2500|
|Collection Enabled:|Yes|
|Auto Collection:|Yes|
|High Water Mark:|80 Percent|
|Max Samples at PC:|69 Days|
|Last Collect Date:|12/5/2022|
|Last Collect Time:|07:00:09|
|Trend On Event:|No|
||

## Result Excel

|Point System Name|Point Name|Panel Name|Descriptor|Point Type|Point Address|Proof Point Address|Engineering Units|COV Limit|Sensor Type|Slope|Intercept|# of Decimal Places|Mode Delay (min)|Level Delay (sec)|Differential|Setpoint Value/Name|Offset1|Priority1|Offset2|Priority2|Mode Point|Trended|RENO Normal|RENO Failed|RENO Pri1|RENO Pri2|RENO Pri3|RENO Pri4|RENO Pri5|RENO Pri6 |
|-|-|-|-|-|-|-|-|-|-|-|-|-|-|-|-|-|-|-|-|-|-|-|-|-|-|-|-|-|-|-|
|Example|Example|PNL1"," (PNL1)"|VIRT Example|LDI|-Virtual-|||||||||||||||||Y|Example Alarming Group|Example Alarming Group|Example Alarming Group|Example Alarming Group|Example Alarming Group|Example Alarming Group|Example Alarming Group|Example Alarming Group|
