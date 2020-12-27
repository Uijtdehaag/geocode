## https://pandas.pydata.org/
## Geocoders:
## https://services.arcgisonline.nl/arcgis/rest/services/Geocoder_BAG/GeocodeServer/findAddressCandidates
## https://services.arcgisonline.nl/arcgis/rest/services/Geocoder_BAG_RD/GeocodeServer/findAddressCandidates
import pandas as pd
import requests, json

def Geocoder_BAG(SingleLineAdres):
	url = 'https://services.arcgisonline.nl/arcgis/rest/services/Geocoder_BAG/GeocodeServer/findAddressCandidates?SingleLine={}&f=pjson'.\
		format(SingleLineAdres)

	ret = requests.get(url)
	if (ret.status_code == 200):
		location = ret.json()['candidates'][0]['location']
		#print (ret.text)
		return [location['x'],location['y']]

	return [0,0]

# Dutch geocoder:
def Geocoder_BAG_RD(SingleLineAdres):
	url = 'https://services.arcgisonline.nl/arcgis/rest/services/Geocoder_BAG_RD/GeocodeServer/findAddressCandidates?SingleLine={}&f=pjson'.\
		format(SingleLineAdres)

	ret = requests.get(url)
	if (ret.status_code == 200):
		location = ret.json()['candidates'][0]['location']
		#print (ret.text)
		return [location['x'],location['y']]

	return [0,0]


df = pd.read_excel('demo.xlsx')
print (df.columns)

for i, row in df.iterrows():
	# Combine columns to single line adres string eg Oostmoer 64, 4709BE Nispen
	singlelineadres = "{} {}{}{}, {} {}".format(row.straat, row.huisnummer,
												row.huisletter, row.toevoeging,
												row.postcode, row.woonplaats		
											).replace ("nan","")
	print (singlelineadres)
	#Call geocoder
	location = Geocoder_BAG(singlelineadres)
	locationRD = Geocoder_BAG_RD(singlelineadres)
	
	# Add new columns with coordinates
	df['lng'] = location[0]
	df['lat'] = location[1]
	df['rd_x'] = locationRD[0]
	df['rd_y'] = locationRD[1]

# save to new excel file
df.to_excel('demo_geocode.xlsx')