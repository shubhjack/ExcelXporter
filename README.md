Feel free to post any suggestions or log any issue and enjoy exporting.

Just create any model list and pass that to the library and it will generate the excel for you in no time.

✅ What's New
🎨 Header styling with background color, font color, and bold text

📐 Cell alignment (Left, Center, Right)

🖋️ Custom font colors for data cells

📦 Optional borders for all cells with configurable color and style

🧱 Fully customizable via a simple StyleOptions model

**--------------------------------------------------------**
Sample call if you have a single data list to export:

add namespace on top using ExcelXporter and use below -

        [HttpGet("exportxls")]
        public IActionResult TestExportExcel()
        {
            // these styles are optional and no need to create and pass if not needed
            // default styles will be applied if not passed
            var styleOptions = new StyleOptions
            {
                HeaderStyle = new HeaderStyle
                {
                    BackgroundColorHex = "4CAF50",  // Header color
                    FontColorHex = "FFFFFF"         // Header font color
                },
                DefaultCellStyle = new ExcelCellStyle
                {
                    FontColorHex = "333333",  // Font color
                    HorizontalAlignment = TextAlignment.Center  // Alignment
                },
                BorderStyle = new BorderStyle
                {
                    ApplyBorders = true,
                    BorderColorHex = "000000", // black
                    Style = BorderStyleValues.Thin
                }
            };

            List<TestModel> objList = new()
            {
                new TestModel ()
                {
                    Id = 1,
                    Name = "John",
                    Email = "john.doe@gmail.com"
                },
                new TestModel
                {
                    Id = 2,
                    Name = "Wick",
                    Email = "john.wick@gmail.com"
                },
            };
            return objList.ExportToExcel("Output", styleOptions);
        }
*--------------------------------------------------------*

*--------------------------------------------------------*
Sample call if you have a multiple data list to export:

add namespace on top using ExcelXporter and use below -

        [HttpGet("exportmultisheetxls")]
        public IActionResult TestExportMultiSheetExcel()
        {
            // these styles are optional and no need to create and pass if not needed
            // default styles will be applied if not passed
            var styleOptions = new StyleOptions
            {
                HeaderStyle = new HeaderStyle
                {
                    BackgroundColorHex = "4CAF50",  // Header background color
                    FontColorHex = "FFFFFF"         // Header font color
                },
                DefaultCellStyle = new ExcelCellStyle
                {
                    FontColorHex = "333333",  // Font color
                    HorizontalAlignment = TextAlignment.Center  // Alignment
                },
                BorderStyle = new BorderStyle
                {
                    ApplyBorders = true,
                    BorderColorHex = "000000", // black
                    Style = BorderStyleValues.SlantDashDot
                }
            };

            List<TestModel> objList = new()
            {
                new TestModel ()
                {
                    Id = 1,
                    Name = "John",
                    Email = "john.doe@gmail.com"
                },
                new TestModel
                {
                    Id = 2,
                    Name = "Wick",
                    Email = "john.wick@gmail.com"
                },
            };

            List<dynamic> objList2 = new();
            objList2.Add(objList);
            objList2.Add(Get().ToList());
            return objList2.ExportToExcelMultipleSheets("Output", styleOptions); 
        }
*--------------------------------------------------------*

Above sample APIs will download excel with single/multiple sheets
