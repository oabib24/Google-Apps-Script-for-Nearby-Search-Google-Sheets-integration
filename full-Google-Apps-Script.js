function getBusinessesByCity() {
  var apiKey = "API_KEY"; // Replace with your API Key
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var inputSheet = ss.getSheetByName("Inputs");
  var outputSheet = ss.getSheetByName("Outputs"); 

  // Keep the header and clear only data below it
  if (outputSheet.getLastRow() > 1) {
    outputSheet.getRange(2, 1, outputSheet.getLastRow() - 1, outputSheet.getLastColumn()).clearContent();
  }

  // Append headers if not present
  if (outputSheet.getLastRow() === 0) {
    var headers = ["Search City", "Search Keyword", "Business Name", "Business Domain", "Primary Category", "Secondary Categories", "Address", "State", "City", "Place ID", "Google Rating", "User Ratings Total"];
    outputSheet.appendRow(headers);
  }

  // Get input data
  var data = inputSheet.getDataRange().getValues(); // Read all data from Inputs sheet
  data.shift(); // Remove headers (assuming row 1 contains column titles)

  data.forEach(function (row) {
    var expectedCity = row[0]; // City in column A
    var expectedState = row[1]; // State in column B
    expectedCity = expectedCity + ", " + expectedState;
    var keyword = row[2]; // Keyword in column B
    var maxResults = row[7] || 30; // Max Results in column C (default to 30 if empty)

    Logger.log(`Fetching results for: ${keyword} in ${expectedCity}`);

    var url = `https://maps.googleapis.com/maps/api/place/textsearch/json?query=${encodeURIComponent(keyword + " in " + expectedCity)}&key=${apiKey}`;
    var results = [];
    var nextPageToken = null;

    do {
      try {
        var response = UrlFetchApp.fetch(url);
        var json = JSON.parse(response.getContentText());

        if (json.status === "OK") {
          results = results.concat(json.results);

          if (json.next_page_token && results.length < maxResults) {
            nextPageToken = json.next_page_token;
            Utilities.sleep(2000);
            url = `https://maps.googleapis.com/maps/api/place/textsearch/json?pagetoken=${nextPageToken}&key=${apiKey}`;
          } else {
            nextPageToken = null;
          }
        } else {
          Logger.log("Error: " + json.status);
          break;
        }

      } catch (error) {
        Logger.log("API Request Failed: " + error.toString());
        break;
      }
    } while (nextPageToken && results.length < maxResults);

    results = results.slice(0, maxResults);

    results.forEach(function (place) {
      var placeId = place.place_id;
      var name = place.name || "N/A";
      var address = place.formatted_address || "N/A";
      var categories = place.types ? place.types.join(", ") : "N/A";
      var primaryCategory = place.types ? place.types[0] : "N/A";
      var secondaryCategories = place.types ? place.types.slice(1).join(", ") : "N/A";
      var rating = place.rating || "N/A";
      var user_ratings_total = place.user_ratings_total || "N/A";

      var detailsUrl = `https://maps.googleapis.com/maps/api/place/details/json?place_id=${placeId}&fields=website,address_components&key=${apiKey}`;

      try {
        var detailsResponse = UrlFetchApp.fetch(detailsUrl);
        var detailsJson = JSON.parse(detailsResponse.getContentText());

        var website = detailsJson.result.website || "N/A";
        var addressComponents = detailsJson.result.address_components || [];

        var city = "N/A", state = "N/A";

        addressComponents.forEach(component => {
          if (component.types.includes("locality")) {
            city = component.long_name;
          }
          if (component.types.includes("administrative_area_level_1")) {
            state = component.short_name;
          }
        });

        if (city.toLowerCase() === expectedCity.toLowerCase().split(",")[0]) {
          outputSheet.appendRow([expectedCity, keyword, name, website, primaryCategory, secondaryCategories, address, state, city, placeId, rating, user_ratings_total]);
        }

      } catch (detailsError) {
        Logger.log("Error fetching details: " + detailsError);
      }
    });

    Logger.log(`Fetched ${results.length} results for "${keyword}" in "${expectedCity}".`);
  });

  Logger.log("All searches completed.");
}
