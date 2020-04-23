# Text analysis application

This application is developed for the 'Intellus Worldwide' conference to get statistical data based on the information retrieved from forums on the Internet (in this case - forums of people having diabetes)


## Program description

All the required data for further analysis must be stored in CSV files (the structure of files is not important but must have the fields like 'Text', 'Topic' and 'Id' - with unique id values)
Moreover, there should be a dictionary that contains all the words needed for the analysis divided into different categories.
This program is able to analyze CSV files by 'text', 'topic', and 'text&topic'.
The output of this program is an excel file with an analytical report and some charts to represent data graphically.


## Run application
To run this app, you should run the script named 'mainWindow.py', which will launch the GUI application.
This program looks like that:

![alt text](main_window.png)

In the first input, a user selects a folder which contains csv files with data.

In the second input, a user must choose an excel dictionary with the words for the further analysis.

In the last input, a user selects a folder where this program will save all the data.

## Example

As an example, please, choose a folder with the CSV file named 'example_data' (this file is in this repository)

Then, select an excel file named 'dictionary_example' (this excel document is also in the repository as well)

Finally, choose any folder to store the resulting output.

A screenshot example before starting an analysis:

![alt text](launching_example.png)

Here, we will provide statistics by texts & topics, so click the relative button, please.

As a result, in the folder destination, which we chose for saving the result, we will get an excel file named 'text_and_post_analysis' and two png pictures with pie charts.

Below is an example of the generated excel document.

![alt text](output_example.png)

This program is highly useful for people who have to gain an idea and statistics about huge sets of data fast and with no need to learn how to operate with any software!


## Further help
If you have any questions or have some suggestions on how to improve this project, do not hesitate to contact me via e-mail antonskripin@gmail.com
