sap.ui.define(["sap/m/MessageToast", "sap/ui/core/Fragment", "xlsx"],
    function (MessageToast, Fragment) {
        'use strict';
        return {
            excelsheetData: [],
            pDialog:null,   
            uploadFile: function (oEvent) {
                var oView = this.getView();
                if (!this.byId("uploadDialogSet")) {
                    Fragment.load({
                        id: "excel_upload",
                        name: "p2pe256.ext.fragment.fragment1",
                        type: "XML",
                        controller: this
                    }).then((oDialog) => {
                        oView.addDependent(oDialog);
                        this.pDialog = oDialog;
                        this.pDialog.open();
                    })
                }
                else {
                    this.byId("uploadDialogSet").open();

                }
            },
            onCloseDialog: function (oEvent) {
                this.pDialog.close();
                this.pDialog.destroy();
            },
            onUploadSet: function (oEvent) {
                // MessageToast.show("File upload invoked.");
                // checking if excel file contains data or not
                if (!this.excelSheetsData.length) {
                    MessageToast.show("Select file to Upload");
                    return;
                }
                var that = this;
                var oSource = oEvent.getSource();
                // creating a promise as the extension api accepts odata call in form of promise only
                var fnAddMessage = function () {
                    return new Promise((fnResolve, fnReject) => {
                         that.callOdata(fnResolve, fnReject);
                    });
                };
                var mParameters = {
                    sActionLabel: oSource.getText() // or "Your custom text" 
                };
                // calling the oData service using extension api
                this.extensionAPI.securedExecution(fnAddMessage, mParameters);
                this.pDialog.close();
                /* TODO:Call to OData */
            },

            onUploadSetComplete: function (oEvent) {
                // getting the UploadSet Control reference
                var oFileUploader = Fragment.byId("excel_upload", "uploadSet");
                // since we will be uploading only 1 file so reading the first file object
                var oFile = oFileUploader.getItems()[0].getFileObject();
                var reader = new FileReader();
                var that = this;
                //  var excelData={};
                var Tobj = {};

                reader.onload = (e) => {
                    // getting the binary excel file content
                    let xlsx_content = e.currentTarget.result;

                    let workbook = XLSX.read(xlsx_content, { type: 'binary' });
                    // here reading only the excel file sheet- Sheet1
                    var excelData = XLSX.utils.sheet_to_row_object_array(workbook.Sheets["Sheet1"]);

                    workbook.SheetNames.forEach(function (sheetName) {
                        // appending the excel file data to the global variable
                        //   excelData = XLSX.utils.sheet_to_row_object_array(workbook.Sheets["sheetName"]);
                        //    that.excelSheetsData = excelData;
                        Tobj = XLSX.utils.sheet_to_row_object_array(workbook.Sheets["Sheet1"]);
                    });
                    this.excelSheetsData = Tobj;
                    console.log("Excel Data", excelData);
                    console.log("Excel Sheets Data", this.excelSheetsData);
                };
                reader.readAsBinaryString(oFile);

                MessageToast.show("Upload Successful");
            },
            onItemRemoved: function (oEvent) {
                this.excelSheetsData = [];
            },

            // helper method to call OData
            callOdata: function (fnResolve, fnReject) {
                //  intializing the message manager for displaying the odata response messages
                var oModel = this.getView().getModel();

                // creating odata payload object for Building entity
                var payload = {};

                this.excelSheetsData.forEach((value, index) => {
                    // setting the payload data
                    payload = {
                        "RecordId"          : value["RecordId"],
                        "RecordType"        : value["RecordType"],
                        "AgreementId"       : value["AgreementId"],
                        "AgreementItem"     : value["AgreementItem"],
                        "AgreementType"     : value["AgreementType"],
                        "Supplier"          : value["Supplier"],
                        "Material"          : value["Material"],
                        "ShortText"         : value["ShortText"],
                        "TargetQty:"        : value["TargetQty"],
                        "ReleaseOrderQty"   : value["ReleaseOrderQty"],
                        "TargetValue"       : value["TargetValue"],
                        "NetPrice"          : value["NetPrice"],
                        "StartDate"         : value["StartDate"],
                        "EndDate"           : value["EndDate"],
                        "AgreementDate"     : value["AgreementDate"],
                        "RecordStatus"      : value["RecordStatus"],
                        "AribaReference"    : value["AribaReference"],
                        "CompanyCode"       : value["CompanyCode"],
                        "PurchasingOrg"     : value["PurchasingOrg"],
                        "PurchasingGroup"   : value["PurchasingGroup"],
                        "Plant"             : value["Plant"],
                        "StorageLocation"   : value["StorageLocation"]

                    };
                    // setting excel file row number for identifying the exact row in case of error or success
                    payload.ExcelRowNumber = (index + 1);
                    // calling the odata service
                    oModel.create("/ContractsTeam", payload, {
                        success: (result) => {
                            console.log(result);
                            var oMessageManager = sap.ui.getCore().getMessageManager();
                            var oMessage = new sap.ui.core.message.Message({
                                message: "File uploaded: " + result.AgreementType,
                                persistent: true, // create message as transition message
                                type: sap.ui.core.MessageType.Success
                            });
                            oMessageManager.addMessages(oMessage);
                            fnResolve();
                        },
                        error: fnReject
                    });
                });
                this.pDialog.destroy();
                
            }


        };
    });