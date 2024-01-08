Feel free to post any suggestions or log any issue and enjoy exporting.

Just create any model list and pass this to the library and it will generate the excel for you in no time.

*--------------------------------------------------------*
Sample call if you have a single data list to export:

add namespace on top using ExcelXporter and use below -
return objList.ExportToExcel("filename");
*--------------------------------------------------------*

*--------------------------------------------------------*
Sample call if you have a multiple data list to export:

add namespace on top using ExcelXporter and use below -

        [HttpGet]
        [Route("[action]")]
        public IActionResult TestExportexcel()
        {
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

            List<TestModel2> objList2 = new()
            {
                new TestModel2 ()
                {
                    Id = 1,
                    Name = "John",
                    Email = "john.doe@gmail.com",
                    Address = "addr1"
                },
                new TestModel2
                {
                    Id = 2,
                    Name = "Wick",
                    Email = "john.wick@gmail.com",
                    Address = "addr2"
                },
                new TestModel2
                {
                    Id = 3,
                    Name = "Constantine",
                    Email = "c.wick@gmail.com",
                    Address = "addr3"
                },
            };
            // This is the important part. You need to create a dynamic list object and add your list of different classes in it.
            List<dynamic> listObj = new()
            {
                objList,
                objList2
            };
            //now call the below extension method on the list created
            return listObj.ExportToExcelMultipleSheets("filename"); 
        }
Above return statement will export multiple data list and will create separate sheet for each list
*--------------------------------------------------------*

Above code will download a excel.
