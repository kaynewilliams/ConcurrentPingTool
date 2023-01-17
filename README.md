The GUI is built using the PySimpleGUI library and the ping functionality is achieved using the subprocess module, which allows the program to run the 'ping' command-line utility. The GUI has two buttons: 'Ping' and 'New Outlook Email'. The 'Ping' button initiates the pinging of the hosts specified in the input field. If the 'Continuously ping' checkbox is checked, the pinging will be done continuously (every second) until the 'Ping' button is pressed again. If the 'Continuously ping' checkbox is not checked, the pinging will be done once, and the results will be displayed in the table. The 'New Outlook Email' button opens a new email in Microsoft Outlook (assuming it is installed on the machine running the program). The program makes use of threads and a queue to allow the GUI to remain responsive while the pinging is being done in the background. This is useful because it allows the user to interact with the program (e.g., change the list of hosts to ping, or stop continuous pinging) while the pinging is being done.
