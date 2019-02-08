Attribute VB_Name = "SevenDayWeatherForecast"
Option Explicit

'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
' Function GetSevenDayWeatherForecast(city_name As String) As Variant()
'
' Created by: https://github.com/mwwbinder
'
' Description: Provides a city's weather forecast for the next 7 days using data from the website
'              https://www.apixu.com/. The free API provides many data points, but this function in
'              its current state only outputs a date, the minimum temperature, and the average
'              temperature for each day.
'
' Parameters: city_name    - The name of a city, "Red Deer" for example.
'
' Return: The function returns a 2d array containing the 7 day forecast. For each of the 7 columns,
'         the first row holds a date, the second row the minimum temperature, and the third row the
'         average temperature. An example would look like:
'
'         2019-02-08 2019-02-09 2019-02-10 2019-02-11 2019-02-12 2019-02-13 2019-02-14
'         6.2        -3.3       -2.6       0.7        -0.3       1          -1.1
'         -4.8       -5.8       -2.8       -2.8       -0.8       -4.8       -2.8
'
'         If the URL for the API is wrong then the function will catch an error. If so then (0, 0)
'         of the array will contain "API url error"
'
' How to Use: You will need to make a free API key by visiting https://www.apixu.com/api.aspx and
'             clicking the button to get started. Once you have the key, paste it as the assigned
'             value for the variable api_key (replace the string <paste your API key here>).
'
'             If you want to change the number of forecasted days you can change the number at the
'             end of the api_url variable where the string has "&days=7". Note that the API only
'             allows values 1 through 10.
'
'             You can change the data points collected and add more data points by changing tokens
'             that are compared to in the Select statement. Once you have your API key you can click
'             Get Started on the website to see an example JSON output. This will allow you to see
'             what other data points are available such as avghumidity.
'
'             Note that if you change the number of days forecasted or change the number of data
'             points collected you will also need to change the bounds of the
'             seven_day_weather_forecast array to accomodate.
'
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
Function GetSevenDayWeatherForecast(city_name As String) As Variant()
    Dim d8 As Date
    Dim avg_temp As Double
    Dim min_temp As Double
    Dim i As Integer
    Dim j As Integer
    Dim http_object As Object
    Dim api_key As String
    Dim api_url As String
    Dim forecast_info() As String
    Dim json_response As String
    Dim weather_forecast_json() As String
    Dim forecasted_day As Variant
    Dim seven_day_weather_forecast(0 To 2, 0 To 6) As Variant
    Dim token As Variant
    
    Set http_object = CreateObject("MSXML2.XMLHTTP")
    api_key = "<paste your API key here>"
    api_url = "http://api.apixu.com/v1/forecast.json?key=" & api_key & "&q=" & city_name & "&days=7"

    http_object.Open "GET", api_url, False
    http_object.Send
    json_response = http_object.ResponseText

    weather_forecast_json = Split(json_response, "[")
    On Error GoTo api_url_error ' If the api_url is wrong we can pass the first Split, but the
                                ' second will throw an error.
        weather_forecast_json = Split(weather_forecast_json(1), "{")
    
    i = 0
    For Each forecasted_day In weather_forecast_json
        forecast_info = Split(forecasted_day, Chr(34)) ' Char 34 is double quote.
        
        j = 0
        For Each token In forecast_info
            Select Case token
                Case Is = "date"
                    d8 = CDate(forecast_info(j + 2))
                    seven_day_weather_forecast(0, i) = d8
                Case Is = "mintemp_c"
                    min_temp = CDbl(Replace(Replace(forecast_info(j + 1), ":", ""), ",", ""))
                    seven_day_weather_forecast(1, i) = min_temp
                Case Is = "avgtemp_c"
                    avg_temp = CDbl(Replace(Replace(forecast_info(j + 1), ":", ""), ",", ""))
                    seven_day_weather_forecast(2, i) = avg_temp
                    i = i + 1
            End Select
            
            j = j + 1
        Next
    Next
        
Done:
    Set http_object = Nothing
    GetSevenDayWeatherForecast = seven_day_weather_forecast
    Exit Function
api_url_error:
    Set http_object = Nothing
    seven_day_weather_forecast(0, 0) = "API url error"
    GetSevenDayWeatherForecast = seven_day_weather_forecast
End Function
