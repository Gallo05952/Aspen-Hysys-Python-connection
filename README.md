# Aspen Hysys-Python connection

 The code provides a method to connect the simulation software Aspen Hysys and Python

I personally find very useful to organize the variables to change and the desired output in Hysys in an internal spreadsheet.

In this way you have to call a single object in Hysys from Python.

This approach, for my experience, is not possible if the variables is a composition. This is because Hysys rad that as a vector. In this case u have to call the stream, then enters in the composition adress and then u can change it using a vector as input
