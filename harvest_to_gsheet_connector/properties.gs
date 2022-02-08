/**
 * Set the properties required for Harvest API connection
 * Only needs to be done as a one-off and is scoped for a user and script.
 * Once this has been run, we'd recommend removing your credentials
 * and commenting out this section afterwards
*/
function setUserProperties() {

     var userProperties = PropertiesService.getUserProperties();

     userProperties.setProperties({
       'HARVEST_ACCESS_TOKEN': '',
       'HARVEST_ACCOUNT_ID': '',
     });

}

