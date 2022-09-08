/** 
 * Helper function to call MS Graph API endpoint
 * using the authorization bearer token scheme
*/
function callMSGraph(endpoint, token, callback) {
  const headers = new Headers();
  const bearer = `Bearer ${token}`;

  headers.append("Authorization", bearer);

  const options = {
    method: "GET",
    headers: headers
  };

  console.log('request made to Graph API at: ' + new Date().toString());

  fetch(endpoint, options)
    .then(response => response.json())
    .then(response => callback(response, endpoint))
    .catch(error => console.log(error));
}

function callMSGraphProfilePhoto(endpoint, token) {
  const headers = new Headers();
  const bearer = `Bearer ${token}`;

  headers.append("Authorization", bearer);

  const options = {
    method: "GET",
    headers: headers
  };

  console.log('request made to Graph API at: ' + new Date().toString());

  return fetch(endpoint, options)
    .then(response => response.blob())
    .then(imageUrl => {
      console.log(imageUrl)
      // Then create a local URL for that image and print it 
      const imageObjectURL = URL.createObjectURL(imageUrl)
      return imageObjectURL
    })
    .catch(error => console.log(error))

}