/**
 * You can copy and paste this code into script editor
 * Make sure to make named path on spreadsheet called: coinRange
 * the cell column for this path should contain all the coin-names as seen in coin url on coingecko
 * View Tutorial: https://www.youtube.com/
 */

function updateCryptoData(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var coinRange = ss.getRangeByName("coinRange");  //pre-created coinRange in spreadsheet
  var coins = coinRange.getValues();
  var stringOfCoins="";
  //append coin ids from spreadsheet into comma seperated string
  for (i = 0; i < coins.length; i++) {
    stringOfCoins = stringOfCoins + coins[i][0] + ",";
  }
  Logger.log(stringOfCoins);
/**Use bulk fetch api call since 429 error code is a rate limit error: CoinGecko limits requests 
 * to 10 calls per second per IP address
 * */
  var allCoinData = IMPORTJSON("https://api.coingecko.com/api/v3/coins/markets?vs_currency=usd&ids="+stringOfCoins,stringOfCoins);
  //Logger.log(Object.keys(allCoinData).length); //for debugging
  var sss = SpreadsheetApp.getActive().getSheetByName('DATA');
  for (i = 0; i < allCoinData.length; i++) {
    Logger.log(allCoinData[i][1].id);
    sss.getRange(i+3,4).setValue(allCoinData[i][1].current_price);
    sss.getRange(i+3,5).setValue(allCoinData[i][1].max_supply);   
    sss.getRange(i+3,7).setValue(allCoinData[i][1].last_updated); 
    sss.getRange(i+3,6).setValue('=image(' + '"'+ allCoinData[i][1].image + '")');  //just for fun :-)
  }
} //updateCryptoData

/**
 * Complete list of attributes available for api.coingecko.com/api/v3/coins/markets
 * {
		"id": "bitcoin",
		"symbol": "btc",
		"name": "Bitcoin",
		"image": "https://assets.coingecko.com/coins/images/1/large/bitcoin.png?1547033579",
		"current_price": 39406,
		"market_cap": 737428343267,
		"market_cap_rank": 1,
		"fully_diluted_valuation": 827521826715,
		"total_volume": 102401839322,
		"high_24h": 42480,
		"low_24h": 35390,
		"price_change_24h": -469.36341888,
		"price_change_percentage_24h": -1.17708,
		"market_cap_change_24h": -8783526211.901611,
		"market_cap_change_percentage_24h": -1.17708,
		"circulating_supply": 18713700,
		"total_supply": 21000000,
		"max_supply": 21000000,
		"ath": 64805,
		"ath_change_percentage": -39.19301,
		"ath_date": "2021-04-14T11:54:46.763Z",
		"atl": 67.81,
		"atl_change_percentage": 58012.93674,
		"atl_date": "2013-07-06T00:00:00.000Z",
		"roi": null,
		"last_updated": "2021-05-20T17:38:11.727Z"
	}
 */

/**
* Imports JSON data to your spreadsheet Ex: IMPORTJSON("http://myapisite.com","city/population")
* @param url URL of your JSON data as string
* @param coins is a comma seperated string containing the ids of the all the coins on the spreadsheet
* @customfunction
*/
function IMPORTJSON(url,coins){
  try{
    var res = UrlFetchApp.fetch(url);   
    var content = res.getContentText();
    var json = JSON.parse(content);
    Logger.log(json[0]); //ensure fetch worked during debugging
    
    if(typeof(json) === "undefined"){
      return "Node Not Available";
    } else if(typeof(json) === "object"){
      //matchup the coins (from the spreadsheet) requested to the prices returned (from Coingecko) since 
      //some coins may be invalid in the result set. 
      //Also the order of results are not guranteed to be the same as the request!
      var tempArr = [];
      var coinArr = coins.split(",");
      for (i = 0; i < coinArr.length; i++) { //loop through all the coin ids
        var match = false;
        for (var obj in json){ //loop through all the results from Coingecko
          if (coinArr[i] == json[obj].id) {
            tempArr.push([obj,json[obj]]); //grab data if coin is found
            match = true;
            break;
          }
        }
        if (!match) {
            tempArr.push([obj,{id: coinArr[i], current_price: 0}]); //set price to 0 if coin was not found
          }
      }
      return tempArr;
    } else if(typeof(json) !== "object") {
      //Logger.log(json);
      return json;
    }
  }
  catch(err){
      return "Error getting data";  
  }
  
}

/**
 * Useful sites to help debug JSON
 * https://www.javainuse.com/jsonpath
 * https://jsonpath.com/
 * https://www.youtube.com/watch?v=fyyShuTKyhs  JSONPath Tutorial
 */
