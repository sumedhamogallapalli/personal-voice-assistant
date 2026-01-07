import win32com.client
import requests

speaker = win32com.client.Dispatch("SAPI.SpVoice")

def weather(query):
    city = ' '.join(query.split('weather at')[1:]).strip()
    print(f"weather at {city}:")
    speaker.speak(f"the weather at {city} is")
    url = "http://api.openweathermap.org/data/2.5/weather?q=" + city + "&appid=b144a088954018157e4e940f2e272d7d"
    weather_data = requests.get(url).json()

    celsius = int(weather_data['main']['temp'] - 273.15)
    fahrenheit = celsius * 9 / 5 + 32
    humidity = weather_data['main']['humidity']
    windspeed = weather_data['wind']['speed']
    clouds = weather_data['weather'][0]['description']

    print(f"The temperature is {celsius} degree celsius and {fahrenheit} degree fahrenheit")
    speaker.speak(f"The temperature is {celsius} degree celsius and {fahrenheit} degree fahrenheit")
    print(f"Humidity is {humidity} %")
    speaker.speak(f"Humidity is {humidity}%")
    print(f"Windspeed is {windspeed} kmph")
    speaker.speak(f"windspeed is {windspeed} kmph")
    print(f"Clouds are {clouds}")
    speaker.speak(f"clouds are {clouds}")


#weather("weather at New York")
