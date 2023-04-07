var IB = IB || {};
var siteUrl = window.location.protocol + "//" + window.location.host + _spPageContextInfo.siteServerRelativeUrl;
console.log("siteUrl ="+siteUrl );
IB.LatestDocuments = (function () {
	var Div = "";
	var RawData = [];
	var GetData = function (Callback) {
console.log("getdata");
		var myQuery='(STS_ListItem_DocumentLibrary) AND (FileType:doc* OR FileType:DOC* OR FileType:xl* OR FileType:XL* OR FileType:ppt* OR FileType:PPT* OR FileType:pdf*)'
		console.log("my query:"+myQuery);
		var Request = {
			request: {
				SortList: {
					'results': [{
						'__metadata': {
							'type': 'Microsoft.SharePoint.Client.Search.Query.Sort'
						},
						'Direction': 1,
						'Property': "LastModifiedTime"
					}]
				},
				Querytext: myQuery,
				RowLimit: 250,
				TrimDuplicates: true
			}
		}

		// Send Request
		$.ajax({
			url: _spPageContextInfo.siteAbsoluteUrl + "/_api/search/postquery",
			type: "POST",
			data: JSON.stringify(Request),
			headers: {
				"accept": "application/json;odata=verbose",
				"content-type": "application/json;odata=verbose",
				"X-RequestDigest": $("#__REQUESTDIGEST").val()
			},
			success: function (data) {
				RawData = data.d.postquery.PrimaryQueryResult.RelevantResults.Table.Rows.results;

				Callback();
			},
			error: function (error) {
				console.log(JSON.stringify(error));
			}
		});
	}

	var DrawResult = function () {
console.log("drawresult");
		$(Div).html('<ul id="DocList"></ul>');

		var i = 0
		$.map(RawData, function (RawEntry) {
			if(i>10)
				return;

			// Get Data
			var Cells = RawEntry.Cells.results;
			var Data = [];
			$.map(Cells, function (Cell) {
				Data[Cell.Key] = Cell.Value;
			});

			if(Data["LinkingUrl"] == null) {
				FilePath= Data["Path"] 
			} else {
				FilePath= Data["LinkingUrl"] 
			}

			if (Data["FileExtension"] != "aaspx"){
				// Draw Line
				$('#DocList').append('' +
					'<a style="display: flex;text-decoration:none" target="_blank" href="' + FilePath+ '">' +
					'<div style="display: inline-block;"><img src="' + GetIconURL(Data["FileExtension"]) + '" style="width:25px;padding:0px 15px 15px 15px;"></div> ' +
					'<div style="display: inline-block; width:100%"> '+
					'<h5>' + Data["Path"].split("/")[Data["Path"].split("/").length - 1] + ' - ' + moment(Data["LastModifiedTime"]).format("DD/MM/YYYY") + '<h5/>' +
					'</div></a>');
				i++;
			}
		});
	}
var GetIconURL = function (Name) {
		var IcoUrl = siteUrl;
		switch (Name) {
			case "docx":
			case "doc":
				IcoUrl += "/Style Library/OLVBrugge/images/word.png";
				break;

			case "xls":
			case "xlsx":
				IcoUrl += "/Style Library/OLVBrugge/images/excel.png";
				break;

			case "ppt":
			case "pptx":
				IcoUrl += "/Style Library/OLVBrugge/images/powerpoint.png";
				break;

			case "pdf":
				IcoUrl += "/Style Library/OLVBrugge/images/pdf.png";
				break;

			case "one":
			case "onenote":
			case "onepkg":
				IcoUrl += "/Style Library/OLVBrugge/images/onenote.png";
				break;

			default:
				IcoUrl += "/Style Library/OLVBrugge/images/other.png";
				break;
		}

		return IcoUrl;
	}

	return {
		Init: function (div) {
			Div = div;
			GetData(function () {
				DrawResult();
			});
		}
	}
})();
