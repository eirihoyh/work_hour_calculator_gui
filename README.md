# work_hour_calculator_gui
An "app" that will automatically register your "workday" localy into an excel file.

This app will make a new sheet every month, write when you started work, ended work, the date and how many hours you have worked.

### Setup to make app
1. Install python
2. Open a cmd
3. pip install pyinstaller
4. Go to where you saved your files
5. pyinstaller.exe --onefile --icon=your_icon.ico name_of_python_file.py
6. When using the "app", remember to have an excel file named "Your_name timer.xlsx" on the same place as your app
7. Start work by writing Your_name and click start
8. End work by writing Your_name and click end

NB! Do NOT close the app before you click end work, or else it will not register your workday.

