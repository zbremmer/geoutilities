#**GeoUtilities**
This simple Excel add-on is designed to provide basic geography tools by utilizing the Google Maps API and the US Federal Communications Commission API.  

**Units:** Latitude and longitude values are returned in degree pairs (48.201561,-104.56818.
Distance measurements are returned in meters for metric and feet for imperia. 

**API Key:** While it is possible to use these functions without a Google Maps API, heavy usage will require an API key to prevent reaching Google’s usage limits.

**Known Issues:**
1.	If user attempts to run these functions on a large number of records simultaneously (for ex., geocoding a large table), Google may throw a query limit error. A better approach would be to use VBA to iterate through the rows to enter the results while pausing the loop for 1 second in between lookups. 

Requires reference to:
-Microsoft XML, v6.0
-Microsoft VBScript Regular Expressions 5.5


#**Functions**

**Geocode(address, Optional apiKey)**

This function uses an address and returns the latitude / longitude values for that location. The return string is formatted as 12.482829,-129.482828 with latitude / longitude in degrees.

Elevation(latitude, longitude, Optional apiKey, Optional units)**

This function uses a latitude / longitude pair to find the elevation at that location. User may specify ‘metric’ or ‘imperial’ units – the default will be metric. Result string units are meters or feet.

**ZipCode(latitude, longitude, apiKey)**

This function uses latitude and longitude to return the zip code for that location. Note: A valid Google Maps API key is required to use this function.

Elevation(latitude, longitude, Optional apiKey, Optional units)**

This function uses a latitude / longitude pair to find the elevation at that location. User may specify ‘metric’ or ‘imperial’ units – the default will be metric. Result string units are meters or feet.

**TransitDistAddr(origin, destination, Optional mode, Optional units)**

This function uses an address for the start and end points to the find distance between locations. Users can specify transit mode (driving, walking, bicycling, or transit) – default is driving. User may specify ‘metric’ or ‘imperial’ units – default will be metric. Result string units are meters or feet. 

Note that this distance is calculated along streets (driving), bike paths (bicycling), bus routes and rail lines (transit) , or a combination thereof. To find the straight line distance between two points, use geoDistAddr(). 

**TransitDistCoord(startLat, startLon, endLat, endLon, Optional mode, Optional units)**

This function uses latitude and longitude values for the start and end points to find the distance between locations. Users can specify transit mode (driving, walking, bicycling, or transit) – default is driving. User may specify ‘metric’ or ‘imperial’ units – default will be metric. Result string units are meters or feet. 

Note that this distance is calculated along streets (driving), bike paths (bicycling), bus routes and rail lines (transit), or a combination thereof. To find the straight line distance between two points, use geoDistCoord(). 

**GeoDistAddr(startAddress, endAddress, Optional units, Optional apiKey)**

This function uses a start and end address to calculate the straight line distance between the points using the Haversine formula. User may specify ‘metric’ or ‘imperial’ units – default will be metric.

**GeoDistCoord(startLat, startLong, endLat, endLong, Optional units)**

This function uses latitude / longitude pairs for the start and end points and returns the straight line distance between the points using the Haversine formula. User may specify ‘metric’ or ‘imperial’ units – default will be metric.

**CountyByCoord(latitude, longitude)**

This function takes the latitude and longitude value of a location and returns the name of the county in which the coordinates fall (US only)

**FIPSByCoord(latitude, longitude)**

This function takes the latitude and longitude value of a location and returns the FIPS code for the county in which the coordinates fall (US only)

**StateByCoord(latitude, longitude)**

This function takes the latitude and longitude value of a location and returns the name of the state in which the coordinates fall (US only)

**CountyByAddr(address)**

This function takes the address of a location and returns the name of the county in which the coordinates fall (US only)

**FIPSByAddr(address)**

This function takes the address of a location and returns the FIPS code for the county in which the coordinates fall (US only)

**StateByAddr(address)**

This function takes the address of a location and returns the name of the state in which the coordinates fall (US only)

**URLEncode(String, Optional SpaceAsPlus)**

This is a utility function used by other functions to encode strings to pass as URL parameters.
