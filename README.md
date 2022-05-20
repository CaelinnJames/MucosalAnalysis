# MucosalAnalysis
Mucosal Analysis R script created and annotated for a non-R user to allow for easy processing of results without too much user interaction.

Script written for R 4.1.3 

Input data for each sample is one excel file with 6 sheets: OscAmp, OscFreq, and 4 breakdown sheets [Osc1:4]

Desired outcomes as specified by user: 
- Caculate tand (G''/G') for Osc amp and save to file
- Plot G'/G'' and y*(%) (Osc amp)
- G* and y*(%) (Osc amp)
- δ(°) and y*(%) (Osc amp)
- tand and y*(%) (Osc amp)
- Plot frequency sweep graph- frequency and δ(°) (Osc freq)
- Caculate σ*(Pa) value at 45 degrees (δ(°)) for each of the 4 breakdown oscilations
- Plot Osc breakdown intercepts as bar chart (value at 45 degrees) - user would like to be able to plot multiple samples at once for this

User wanted a script that was easy to use for a non-R user that required minimal user interaction to run for multiple samples. For each graph, user wanted a version with axes scaled for the data for that sample, as well as a version with axes scaled for data across all samples to allow for comparison. 

Annotation has been provided throughout the script to explain to the user what to change if they would like to change the output.
