# Creating-WIndows-Service-In-Python

Over the years I found very few resources that assisted me in creating a windows service in Python. An article by [ThePythonCorner](https://www.thepythoncorner.com/2018/08/how-to-create-a-windows-service-in-python/) has assisted me in creating a fully functional windows service.
In this demo I have created a simple windows services to copy all excel files from one folder to another and covert them to csv.

#### Update the service name & description as well as source,destination & logger path  first in app.py
Note : Do not put spaces or special characters on the service name and description
```python
 _svc_name_ = "ServiceName"
 _svc_display_name_ = "ServiceName"
 _svc_description_ = "ServiceDescription"
```

```python
 source = "Source_path"
 destination = "Destination_path"
```

```python
file_handler = logging.FileHandler('file_path_for_logger')
```

#### To install,update and debug the service
To Install the service Open cmd or shell as an administrator, navigate to the root folder of the project(navigate to the folder containing install.py and SMSWinService) and run the command below:
```python
python app.py install
```

In the future if you want to update the code of the windows service go to the root folder again and run:
```python
python app.py update
```

To run the service in DEBUG mode go to the root folder again and run:
```python
python app.py debug
```

#####If the service is not working : [Reference](https://thepythoncorner.com/posts/2018-08-01-how-to-create-a-windows-service-in-python/)
>There are a couple of known problems that can happen writing Windows Services in Python. If you have successfully installed the service but starting it you get an error, follow this iter to troubleshoot your service:


>Check if Python is in your PATH variable. It MUST be there. To check this, just open a command prompt and try starting the python interpreter by typing “python”. If it starts, you are ok.

>Be sure to have the file: (Note assuming you have python3.6 installed on your pc) 

>C:\Program Files\Python36\Lib\site-packages\win32\pywintypes36.dll (please note that “36” is the version of your Python installation). If you don’t have this file, take it from C:\Program Files\Python36\Lib\site-packages\pywin32_system32\pywintypes36.dll and copy it into C:\Program Files\Python36\Lib\site-packages\win32

>If you still have problems, try executing your Python script in debug mode. To try this with our previous example, open a terminal, navigate to the directory where the script resides and type

```python
python app.py debug
```


