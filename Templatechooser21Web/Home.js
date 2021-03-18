
(function () {

    // write the ajax call for get the data
    //make HTTPS request to the server

    //try the ajax call to the server using a JSON ot jquery

    // The initialize function must be run each time a new page is loaded.
    Office.initialize = function (reason) {
        $(document).ready(function () {
            //scrollMenu();
            //cardSelection();
            displayImage();
            // If not using Word 2016, use fallback logic.
            if (!Office.context.requirements.isSetSupported('WordApi', '1.1')) {
                console.log('Sorry. The tutorial add-in uses Word.js APIs that are not available in your version of Office.');
            }

            // Do something that is only available via the new APIs
            //Selection of image, insert image
            $('.imgchooser').on('click', function (event) {
                var image = event.currentTarget.querySelector("image");
                var src = image.src;

                //to insert the image from the function!
                toDataURL(src, function (dataUrl) {
                    insertImages(dataUrl);
                });
            });
            //function 2 to inser the image
            $('.imgchooser2').on('click', function (event) {
                var image = event.currentTarget.querySelector("img");
                var src = image.src;
                console.log("insert image2");
                //if this the image of the Image

                //to insert the image from the function!
                toDataURL(src, function (dataUrl) {
                    insertImages(dataUrl);

                });
            });

        });

    };

    //Function for display the imagine in the addin
    function displayimage() {
        //array of Strings of objects to display the image.
        var Technology = ["human-brain.jpg", "cancun.jpg", "Montreal123.jpg", "aurora.jpg", "Calgary.jpg", "altstadt.jpg", "cancun.jpg", "blackhole.jpg", "Sunsetbeach.jpg",
            "freedownload.jpg", "photocopy.jpg", "Montreal123.jpg", "aurora.jpg", "altstadt.jpg", "Calgary.jpg", "altstadt.jpg", "cancun.jpg", "101_2018_1_18.jpg", "101_2018_1_19.jpg", "101_2018_3_20.jpg"];
        //for look for the image.
        for (var i = 0; i < Technology.length; i++) {
            var image = Technology[i];
            //add-in container for display the imagine with the url and the class html addin 
            $('.myimage-container').append(
                '<div class="imgchooser">' +
                '<tr id="Technology-Equipment"><td><img src = "Images/cancun.jpg' + image + '"style="width:100%" height="100%" "align="right" "" alt = "" "class ="filterDiv " ></td></tr> ' +
                '</div>'
            );
            //forlook for the image.
            $(".imgchooser").show();
        }
        console.log("images to search it");
    }
    // function DisplayImage

    function displayImage() {
        var request = new XMLHttpRequest();
        request.open("GET", "phpmethod");
        request.onreadystatechange = function () {
            // check if (the request is compete and was successful)
            if (this.readyState === 4 && this.status === 200) {
                document.getElementById("result").innerHTML = this.responseText;
            }
        };
        request.send();

    }

    // toDataUrl fuction to get the date of the image
    function toDataURL(url, callback) {
        url = "https://srk.sharepoint.com/sites/CommunicationsNA/ImageChooserTest/Forms/AllItems.aspx";
        //method for the request of the data
        var xhr = new XMLHttpRequest();
        xhr.onload = function () {
            var reader = new FileReader();
            reader.onloadend = function () {
                callback(reader.result.split(',')[1]);
                callback(getSelection(insertImages))
            }
            reader.readAsDataURL(xhr.response);
        };
        //to open the url and get the data of the image selected
        xhr.open('GET', url);
        xhr.responseType = 'blob';
        xhr.send();
        console.log('toDataURL');
    }
    // the function to get the database of the image.
    function insertImages(base64) {
        Word.run(function (context) {
            // Queue a command to get the current selection.
            // Create a proxy range object for the selection.
            var range = context.document.getSelection();
            // Queue a command to replace the selected text.
            range.insertInlinePictureFromBase64(base64, Word.InsertLocation.replace);
            // Synchronize the document state by executing the queued commands,
            // and return a promise to indicate task completion.
            return context.sync().then(function () {
                console.log('Added an image.');
            });
        })
            .catch(function (error) {
                console.log('Error: ' + JSON.stringify(error));
                if (error instanceof OfficeExtension.Error) {
                    console.log('Debug info: ' + JSON.stringify(error.debugInfo));
                }
            });
    }

})();
