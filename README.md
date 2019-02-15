# VBA
This script takes 7 column spreadsheet with columns: Ticker, date, open price, high price, low price, close price, stock volume.

Below is an example of an input file

Input:
<ticker>	<date>	<open>	<high>	<low>	<close>	<vol>
A	20160101	41.81	-42.36365509	41.81	41.81	0
A	20160104	41.06	-1.305047274	40.34	40.69	3287300
A	20160105	40.73	0.902213812	40.34	40.55	2587200
A	20160106	40.24	0.71801573	40.05	40.73	2103600

Below is the 2 outputs that are produced by the script on the same sheet as the input data.

Output:
Ticker	Yearly Change	  Percent Change	Total Stock Volume
P	      -8.77	          -32.97%	        2,087,326,500
PAA	    -0.45	          -0.87%	        265,764,400
PAC	    9.98	          18.75%	        15,869,800
PACD	  -68.2	          -59.51%	        143,918,500
PAG	    1.91	          4.05%	          88,908,800


Output:
	                    Ticker	Value
Greatest % Increase	  PANW	113.28%
Greatest % Decrease	  PWE	  -75.12%
Greatest Total Volume	PBR	  7,733,747,800

