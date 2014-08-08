$(document).ready(function () {

    $("#btnSaveProduct").click(function (e) {
        e.preventDefault();

        var spHostUrl = getSPHostUrlFromQueryString(window.location.search);

        var urlAddProduct = "/Home/AddProduct?SPHostUrl=" + spHostUrl;

        $.post(urlAddProduct,
                {
                    title: $("#productTitle").val(),
                    description: $("#productDescription").val(),
                    price: $("#productPrice").val(),
                }).done(function () {
                    $("#myModal").modal('hide');
                    location.reload();
                })
                .fail(function () {
                    alert("Failed to add the new product!");
                });
    });

    // Gets SPHostUrl from the given query string.
    function getSPHostUrlFromQueryString(queryString) {
        if (queryString) {
            if (queryString[0] === "?") {
                queryString = queryString.substring(1);
            }

            var keyValuePairArray = queryString.split("&");

            for (var i = 0; i < keyValuePairArray.length; i++) {
                var currentKeyValuePair = keyValuePairArray[i].split("=");

                if (currentKeyValuePair.length > 1 && currentKeyValuePair[0] == "SPHostUrl") {
                    return currentKeyValuePair[1];
                }
            }
        }

        return null;
    }

});