# simplyExcelFunction

Description : Personal Excel function collection that support me work faster in excel 

## Getting Started 
Build the package and copy the 64bit `.xll` from `bin\Debug\net10.0-windows\publish\SimpleExcelfunctions-AddIns64.xll` to your folder in your PC. 

## Add into Excel 
1. Open Excel
2. Open the Options from the File menu item 
3. in the Add-ins page , find Manage = Excel Add-ins. 
    <div style="text-align: center;">
    <img src="readme\1.png" alt="Centered image">
    </div>
4. Click Go 
5. use Browse to add your `.xll` into your excel. 
6. Restart the Excel

## Build the `xll` or addon file 
```
dotnet clean && dotnet build -c Release
```

## Run The test case 
```
dotnet clean && dotnet restore --force && dotnet test --no-restore --logger "console;verbosity=normal"
```
## Current Function 
**`=reDateTime($cell)`** use to convert a string date time format(yyyy-mm-dd hh:mm:ss) to excel date time format. 
**`=unixToDateTime($cell)`** convert the unix time to Local Date Time
**`=dateTimeToUnix($cell)`** convert the Local Date Time to unix time