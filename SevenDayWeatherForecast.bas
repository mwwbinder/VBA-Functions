Attribute VB_Name = "SevenDayWeatherForecast"
Option Explicit

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
    api_key = "76e8ca04d2c04561829153509190702"
    api_url = "http://api.apixu.com/v1/forecast.json?key=" & api_key & "&q=" & city_name & "&days=7"

    http_object.Open "GET", api_url, False
    http_object.Send
    json_response = http_object.ResponseText

    weather_forecast_json = Split(json_response, "[")
    On Error GoTo api_url_error
        weather_forecast_json = Split(weather_forecast_json(1), "{")
    
    i = 0
    For Each forecasted_day In weather_forecast_json
        forecast_info = Split(forecasted_day, Chr(34))
        
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
    GetSevenDayWeatherForecast = seven_day_weather_forecast
    Exit Function
api_url_error:
    seven_day_weather_forecast(0, 0) = "API url error"
    GetSevenDayWeatherForecast = seven_day_weather_forecast
End Function
