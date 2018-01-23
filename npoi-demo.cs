        [HttpPost]
        public string Import()
        {
            HttpPostedFileBase file = Request.Files["file"];
            if (file == null)
            {
                return JsonConvert.SerializeObject(new { Success = false, Msg = "文件为空" });
            }
            //检查文件类型
            string filename = Path.GetFileName(file.FileName);
            string ext = Path.GetExtension(filename);

            if (string.IsNullOrEmpty(ext) || !(ext.TrimStart('.').Equals("xls", StringComparison.OrdinalIgnoreCase) || ext.TrimStart('.').Equals("xlsx", StringComparison.OrdinalIgnoreCase)))
            {
                return JsonConvert.SerializeObject(new { Success = false, Msg = "非法的Excel类型" });
            }
            ext = ext.TrimStart('.').ToLower();
            string type = "";
            string path = "/upload/";
            type = "file";
            path = Path.Combine(path, "file");
            try
            {
                filename = MD5Helper.GetStreamMD5(file.InputStream); //使用文件的md5值作为文件名，相同文件直接覆盖存储   
                string mapPath = Server.MapPath(path);
                if (!Directory.Exists(mapPath))
                {
                    Directory.CreateDirectory(mapPath);
                }
                file.SaveAs(Path.Combine(mapPath, filename + "." + ext));
                List<string> sqls = new List<string>();
                Dictionary<int, string> headerDictionary = new Dictionary<int, string>();
                IWorkbook workbook = null;
                if (ext.Equals("xls", StringComparison.OrdinalIgnoreCase))
                {
                    workbook = new HSSFWorkbook(System.IO.File.OpenRead(Server.MapPath(path + "/" + filename + "." + ext)));
                }
                else
                {
                    workbook = new XSSFWorkbook(Server.MapPath(path + "/" + filename + "." + ext));
                }
                //获取第一个表
                ISheet sheet = workbook.GetSheetAt(0);
                //获取行 默认认为第一行为表头
                if (sheet.LastRowNum > 0)
                {
                    IRow header = sheet.GetRow(0);
                    for (int j = 0; j < header.LastCellNum; j++)
                    {
                        ICell cell = header.GetCell(j);
                        if (cell != null)
                        {
                            headerDictionary.Add(j, cell.ToString().Trim());
                        }
                    }
                    for (int i = 1; i <= sheet.LastRowNum; i++)
                    {
                        IRow row = sheet.GetRow(i);
                        if (row != null)
                        {
                            string name = "";
                            string nickname = "";
                            foreach (KeyValuePair<int, string> pair in headerDictionary)
                            {
                                if (pair.Value.Equals("姓名"))
                                {
                                    ICell cell = row.GetCell(pair.Key);
                                    if (cell != null)
                                    {
                                        name = cell.ToString().Trim();
                                    }
                                }
                                else if (pair.Value.Equals("昵称"))
                                {
                                    ICell cell = row.GetCell(pair.Key);
                                    if (cell != null)
                                    {
                                        nickname = cell.ToString().Trim();
                                    }
                                }
                            }
                            if (!string.IsNullOrEmpty(name) && !string.IsNullOrEmpty(nickname))
                            {
                                sqls.Add("insert into userinfo(userid,nickname,pwd) values('" + name + "','" + nickname + "','e10adc3949ba59abbe56e057f20f883e');");
                            }
                        }
                    }
                }
                if (sqls.Count == 0)
                {
                    return JsonConvert.SerializeObject(new { Success = false, Msg = "文件中数据为空" });
                }
                if (_userinfoBll.InsertSqls(sqls) > 0)
                {
                    return JsonConvert.SerializeObject(new { Success = true, @Type = type, Msg = "导入成功", Data = path + "/" + filename + "." + ext, FileName = filename + "." + ext });
                }
                return JsonConvert.SerializeObject(new { Success = false, Msg = "导入失败" });
            }
            catch (Exception e)
            {
                return JsonConvert.SerializeObject(new { Success = false, Msg = "导入失败" });
            }

        }

        public FileResult Export()
        {
            HSSFWorkbook workbook = new HSSFWorkbook();
            //添加一个sheet
            NPOI.SS.UserModel.ISheet sheet = workbook.CreateSheet("用户信息");
            System.IO.MemoryStream ms = new System.IO.MemoryStream();

            //给sheet1添加第一行的头部标题
            NPOI.SS.UserModel.IRow row = sheet.CreateRow(0);

            ICellStyle style = workbook.CreateCellStyle();
            style.Alignment = HorizontalAlignment.Center;
            style.WrapText = true;
            IFont font = workbook.CreateFont();
            font.FontHeightInPoints = 11;
            font.Boldweight = (short)NPOI.SS.UserModel.FontBoldWeight.Bold;
            font.FontName = "微软雅黑";
            style.SetFont(font);//HEAD 样式

            row.CreateCell(0).SetCellValue("姓名");
            row.CreateCell(1).SetCellValue("昵称");
            row.CreateCell(2).SetCellValue("添加时间");

            row.Cells[0].CellStyle = style;
            row.Cells[1].CellStyle = style;
            row.Cells[2].CellStyle = style;
            List<AdminInfo> userinfos = _userinfoBll.GetModelList(" 1=1 order by createtime desc");
            for (int j = 0; j < userinfos.Count; j++)
            {
                NPOI.SS.UserModel.IRow rowtemp = sheet.CreateRow(j + 1);
                rowtemp.CreateCell(0).SetCellValue(userinfos[j].userid);
                rowtemp.CreateCell(1).SetCellValue(userinfos[j].nickname);
                rowtemp.CreateCell(2).SetCellValue(userinfos[j].createtime.ToString("yyyy-MM-dd HH:mm:ss"));
            }

            workbook.Write(ms);
            ms.Seek(0, SeekOrigin.Begin);
            string fileName = "用户信息" + DateTime.Now.ToString("yyMMddHHmmssfff") + ".xls";
            return File(ms, "application/vnd.ms-excel", fileName);
        }
