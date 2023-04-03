var IB = IB || {};

var siteUrl = window.location.protocol + "//" + window.location.host + _spPageContextInfo.siteServerRelativeUrl;
IB.LatestDocuments = (function () {
	var Div = "";
	var RawData = [];

	var GetData = function (Callback) {

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
				Querytext: '(path:'+siteUrl+'/") contentclass:"ǂǂ5354535f4c6973744974656d5f446f63756d656e744c696272617279"',
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

		$(Div).html('<ul id="DocList"></ul>');

		var i = 0
		$.map(RawData, function (RawEntry) {
			if(i>25)
				return;

			// Get Data
			var Cells = RawEntry.Cells.results;
			var Data = [];
			$.map(Cells, function (Cell) {
				Data[Cell.Key] = Cell.Value;
			});

			if (Data["FileExtension"] != "aspx" && Data["OriginalPath"].toLowerCase().indexOf("/ds/") === -1 && Data["OriginalPath"].toLowerCase().indexOf("/amicale/") === -1){
				// Draw Line
				$('#DocList').append('' +
					'<a style="display: flex;text-decoration:none" target="_blank" href="' + Data["Path"] + '">' +
					'<div style="display: inline-block; width:100%"> <h4 style="width:99%">' + Data["Title"] + '</h4>' +
					'<h5>' + Data["Path"].split("/")[Data["Path"].split("/").length - 1] + ' - ' + moment(Data["LastModifiedTime"]).format("DD/MM/YYYY") + '<h5/>' +
					//'<h6 style="color:#aaa;font-size:0.6em">' + Data["Author"] + '<h6/>' +
					'</div></a>');
				i++;
			}
		});
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
