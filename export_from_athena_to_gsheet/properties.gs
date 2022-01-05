
/**
 * Set the properties required for AWS connection or Athena export
 * Only needs to be done as a one-off and is scoped for a user and script.
 * Once this has been run, we'd recommend removing your credentials
 * and commenting out this section afterwards
*/
function setUserProperties() {

     var userProperties = PropertiesService.getUserProperties();

     userProperties.setProperties({
       'AWS_ACCESS_KEY': '',
       'AWS_SECRET_KEY': '',
       'AWS_REGION': '',
       'ATHENA_S3_OUTPUT_LOCATION': ''
     });

}
