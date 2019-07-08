using CMArchivoSharePoint.Model;
using LinqToExcel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Web;
using Microsoft.SharePoint.Client;
using SP = Microsoft.SharePoint.Client;
using System.Data;

namespace CMArchivoSharePoint
{
    class Program
    {
        static void Main(string[] args)
        {
            string pathArchivoCompleto = "C:\\Users\\k697344\\Documents\\Comex PPG\\Documentacion\\Documentos Control Documental_V1_004.xlsx";
            string nombrePestana = "CM FINAL";


            string usuarioSharePoint = "S004221";
            string passwordSharePoint = "Julio2019";

            string urlShareFolder = "\\\\10.104.175.150\\Campania\\Reporte";
            string urlCompletoFolder = "";

            string urlSharePointOrigen = "https://one.web.ppg.com/la/mexico/ppgmexico/CalidadTotal/Control_Documental/BckP_SJDR";
            string urlSharePointDestino = "https://one.web.ppg.com/la/mexico/ppgmexico/CalidadTotal/Control_Documental/DocsPublic";
            string urlCompletoOrigen = "";
            string urlCompletoDestino = "";

            string siteUrl = "https://one.web.ppg.com/la/mexico/ppgmexico/CalidadTotal/Control_Documental/";
            string bibliotecaDocumentoSP = "Manual de Calidad";
            string catalogoArea = "Area";
            string catalogoDepartamento = "Department";
            string catalogoDocType = "DocType";
            string catalogoSBU = "SBU";
            string catalogoCliente = "Clientes";

            ExcelQueryFactory book = new ExcelQueryFactory();
            List<Archivo> ListArchivo = new List<Archivo>();

            List<Archivo> ListArchivoEncontrado = new List<Archivo>();

            //string nombreArchivo = "\\\\10.104.175.150\\Campania\\Reporte\\Guid.NewGuid().ToString()" + ".xls";
            string pathArchivoExcel = "\\\\10.104.175.150\\Campania\\Reporte\\Archivos_Cargados_" + Guid.NewGuid().ToString() + ".xls";

            DataSet dsArchivoExcel = new DataSet();
            DataTable dtArchivoExcel = new DataTable();

            DataSet dsCatalogos = new DataSet();
            DataTable dtArea = new DataTable();
            DataTable dtDepartamento = new DataTable();
            DataTable dtDocType = new DataTable();

            dtArchivoExcel.Columns.Add("Area");
            dtArchivoExcel.Columns.Add("Departamento");
            dtArchivoExcel.Columns.Add("TipoDocumento");
            dtArchivoExcel.Columns.Add("DepartamentoCodigo");
            dtArchivoExcel.Columns.Add("Codigo");
            dtArchivoExcel.Columns.Add("NombreDocumento");
            dtArchivoExcel.Columns.Add("DescripcionDocumento");
            dtArchivoExcel.Columns.Add("NumeroRevision");
            dtArchivoExcel.Columns.Add("FCambioFijo");
            dtArchivoExcel.Columns.Add("FCambioFrecuente");
            dtArchivoExcel.Columns.Add("SBU");
            dtArchivoExcel.Columns.Add("Cliente");

            try
            {
                ////LECTURA DEL ARCHIVO EXCEL
                //book = new ExcelQueryFactory(pathArchivoCompleto);

                //ListArchivo = book.Worksheet(nombrePestana).AsEnumerable()
                //                .Select(n => new Archivo
                //                {
                //                    Area = n["Area"].Cast<string>(),
                //                    Departamento = n["Department"].Cast<string>(),
                //                    TipoDocumento = n["Document"].Cast<string>(),
                //                    DepartamentoCodigo = n["Department Code"].Cast<string>(),
                //                    Codigo = n["Archivo"].Cast<string>(),
                //                    NombreDocumento = n["Archivo"].Cast<string>(),
                //                    DescripcionDocumento = n["Name Document"].Cast<string>(),
                //                    NumeroRevision = n["Revision"].Cast<string>(),
                //                    FCambioFijo = n["Date Revision"].Cast<string>(),
                //                    FCambioFrecuente = n["Date Revision"].Cast<string>(),
                //                    SBU = n["SBU"].Cast<string>(),
                //                    Cliente = n["Cliente"].Cast<string>()
                //                }).ToList();

                ////CARGAR CLIENTE
                ////ListArchivo = ListArchivo.Where(n => !string.IsNullOrEmpty(n.DescripcionDocumento)).ToList();

                //ListArchivo = ListArchivo.Where(n => !string.IsNullOrEmpty(n.Area) && n.Area.ToUpper() == "SATELITES").ToList();

                #region CARGAR ORIGEN DESTINO

                /*
                //CARGAR ARCHIVOS A SHARE POINT
                foreach (Archivo a in ListArchivo.ToList())
                    //.Where(n => !string.IsNullOrEmpty(n.Codigo) && n.Codigo.Trim() == "IT-1509").ToList())
                {
                    if (!string.IsNullOrEmpty(a.Codigo))
                    {
                        //SE BUSCA EN FORMATO EXCEL
                        ///////////////////////////
                        urlCompletoOrigen = Path.Combine(HttpUtility.HtmlEncode(urlSharePointOrigen), a.Codigo.Trim()); //+ ".xls");
                        urlCompletoFolder = Path.Combine(HttpUtility.HtmlEncode(urlShareFolder), a.Codigo.Trim()); //+ ".xls");
                        urlCompletoDestino = Path.Combine(HttpUtility.HtmlEncode(urlSharePointDestino), a.Codigo.Trim()); //+ ".xls");

                        try
                        {
                            using (WebClient client = new WebClient())
                            {
                                client.Credentials = new NetworkCredential(usuarioSharePoint, passwordSharePoint);
                                client.DownloadFile(urlCompletoOrigen, urlCompletoFolder);
                                client.UploadFile(urlCompletoDestino, "PUT", urlCompletoFolder);

                                a.Codigo = a.Codigo.Trim(); //+ ".xls";

                                ListArchivoEncontrado.Add(a);

                                continue;
                            }

                            //a.Codigo = a.Codigo.Trim(); //+ ".xls";

                            //ListArchivoEncontrado.Add(a);
                        }
                        catch (Exception ex)
                        {
                            continue;
                        }

                        ////INTENTO SIN ESPACIOS EN BLANCO
                        //urlCompletoOrigen = Path.Combine(HttpUtility.HtmlEncode(urlSharePointOrigen), a.Codigo.Trim().Replace(" ", "") + ".xls");
                        //urlCompletoFolder = Path.Combine(HttpUtility.HtmlEncode(urlShareFolder), a.Codigo.Trim().Replace(" ", "") + ".xls");
                        //urlCompletoDestino = Path.Combine(HttpUtility.HtmlEncode(urlSharePointDestino), a.Codigo.Trim().Replace(" ", "") + ".xls");

                        //try
                        //{
                        //    using (WebClient client = new WebClient())
                        //    {
                        //        client.Credentials = new NetworkCredential(usuarioSharePoint, passwordSharePoint);
                        //        client.DownloadFile(urlCompletoOrigen, urlCompletoFolder);
                        //        client.UploadFile(urlCompletoDestino, "PUT", urlCompletoFolder);

                        //        a.Codigo = a.Codigo.Trim().Replace(" ", "") + ".xls";

                        //        ListArchivoEncontrado.Add(a);

                        //        continue;
                        //    }
                        //}
                        //catch (Exception ex)
                        //{

                        //}

                        ////SE BUSCA EN FORMATO WORD
                        ////////////////////////////
                        //urlCompletoOrigen = Path.Combine(HttpUtility.HtmlEncode(urlSharePointOrigen), a.Codigo.Trim() + ".doc");
                        //urlCompletoFolder = Path.Combine(HttpUtility.HtmlEncode(urlShareFolder), a.Codigo.Trim() + ".doc");
                        //urlCompletoDestino = Path.Combine(HttpUtility.HtmlEncode(urlSharePointDestino), a.Codigo.Trim() + ".doc");

                        //try
                        //{
                        //    using (WebClient client = new WebClient())
                        //    {
                        //        client.Credentials = new NetworkCredential(usuarioSharePoint, passwordSharePoint);
                        //        client.DownloadFile(urlCompletoOrigen, urlCompletoFolder);
                        //        client.UploadFile(urlCompletoDestino, "PUT", urlCompletoFolder);

                        //        a.Codigo = a.Codigo.Trim().Replace(" ", "") + ".doc";

                        //        ListArchivoEncontrado.Add(a);

                        //        continue;
                        //    }
                        //}
                        //catch (Exception ex)
                        //{

                        //}

                        ////INTENTO SIN ESPACIOS EN BLANCO
                        //urlCompletoOrigen = Path.Combine(HttpUtility.HtmlEncode(urlSharePointOrigen), a.Codigo.Trim().Replace(" ", "") + ".doc");
                        //urlCompletoFolder = Path.Combine(HttpUtility.HtmlEncode(urlShareFolder), a.Codigo.Trim().Replace(" ", "") + ".doc");
                        //urlCompletoDestino = Path.Combine(HttpUtility.HtmlEncode(urlSharePointDestino), a.Codigo.Trim().Replace(" ", "") + ".doc");

                        //try
                        //{
                        //    using (WebClient client = new WebClient())
                        //    {
                        //        client.Credentials = new NetworkCredential(usuarioSharePoint, passwordSharePoint);
                        //        client.DownloadFile(urlCompletoOrigen, urlCompletoFolder);
                        //        client.UploadFile(urlCompletoDestino, "PUT", urlCompletoFolder);

                        //        a.Codigo = a.Codigo.Trim().Replace(" ", "") + ".doc";

                        //        ListArchivoEncontrado.Add(a);

                        //        continue;
                        //    }
                        //}
                        //catch (Exception ex)
                        //{

                        //}

                        ////SE BUSCA EN FORMATO PDF
                        ///////////////////////////
                        //urlCompletoOrigen = Path.Combine(HttpUtility.HtmlEncode(urlSharePointOrigen), a.Codigo.Trim() + ".pdf");
                        //urlCompletoFolder = Path.Combine(HttpUtility.HtmlEncode(urlShareFolder), a.Codigo.Trim() + ".pdf");
                        //urlCompletoDestino = Path.Combine(HttpUtility.HtmlEncode(urlSharePointDestino), a.Codigo.Trim() + ".pdf");

                        //try
                        //{
                        //    using (WebClient client = new WebClient())
                        //    {
                        //        client.Credentials = new NetworkCredential(usuarioSharePoint, passwordSharePoint);
                        //        client.DownloadFile(urlCompletoOrigen, urlCompletoFolder);
                        //        client.UploadFile(urlCompletoDestino, "PUT", urlCompletoFolder);

                        //        a.Codigo = a.Codigo.Trim() + ".pdf";

                        //        ListArchivoEncontrado.Add(a);

                        //        continue;
                        //    }
                        //}
                        //catch (Exception ex)
                        //{

                        //}

                        ////INTENTO SIN ESPACIOS EN BLANCO
                        //urlCompletoOrigen = Path.Combine(HttpUtility.HtmlEncode(urlSharePointOrigen), a.Codigo.Trim().Replace(" ", "") + ".pdf");
                        //urlCompletoFolder = Path.Combine(HttpUtility.HtmlEncode(urlShareFolder), a.Codigo.Trim().Replace(" ", "") + ".pdf");
                        //urlCompletoDestino = Path.Combine(HttpUtility.HtmlEncode(urlSharePointDestino), a.Codigo.Trim().Replace(" ", "") + ".pdf");

                        //try
                        //{
                        //    using (WebClient client = new WebClient())
                        //    {
                        //        client.Credentials = new NetworkCredential(usuarioSharePoint, passwordSharePoint);
                        //        client.DownloadFile(urlCompletoOrigen, urlCompletoFolder);
                        //        client.UploadFile(urlCompletoDestino, "PUT", urlCompletoFolder);

                        //        a.Codigo = a.Codigo.Trim().Replace(" ", "") + ".pdf";

                        //        ListArchivoEncontrado.Add(a);

                        //        continue;
                        //    }
                        //}
                        //catch (Exception ex)
                        //{

                        //}

                    }
                }

                */
                #endregion


                ClientContext clientContext = new ClientContext(siteUrl);
                SP.Web myWeb = clientContext.Web;
                List myListArchivos = myWeb.Lists.GetByTitle(bibliotecaDocumentoSP);

                ListItemCollection listItems = myListArchivos.GetItems(CamlQuery.CreateAllItemsQuery());
                clientContext.Load(listItems);
                clientContext.ExecuteQuery();


                #region CREAR ARCHIVO EXCEL

                /////////////////////
                //////CREAR ARCHIVO EXCEL

                //listItems.ToList().ForEach(item =>
                //{
                //    dtArchivoExcel.Rows.Add(item["Area"] == null ? "" : ((FieldLookupValue)item["Area"]).LookupValue,
                //                                    item["Department"] == null ? "" : ((FieldLookupValue)item["Department"]).LookupValue,
                //                                    item["DocType"] == null ? "" : ((FieldLookupValue)item["DocType"]).LookupValue,
                //                                    item["DepartmentCode"] == null ? "" : item["DepartmentCode"].ToString(),
                //                                    item["FileLeafRef"] == null ? "" : item["FileLeafRef"].ToString(),
                //                                    item["Title"] == null ? "" : item["Title"].ToString(),
                //                                    item["Cliente"] == null ? "" : ((FieldLookupValue)item["Cliente"]).LookupValue,
                //                                    item["Revision"] == null ? "" : item["Revision"].ToString(),
                //                                    item["Update"] == null ? "" : item["Update"].ToString(),
                //                                    item["Created"] == null ? "" : item["Created"].ToString(),
                //                                    item["SBU"] == null ? "" : ((FieldLookupValue)item["SBU"]).LookupValue,
                //                                    item["Modified"] == null ? "" : item["Modified"].ToString());
                //});


                //dsArchivoExcel.Tables.Add(dtArchivoExcel);

                //ExcelLibrary.DataSetHelper.CreateWorkbook(pathArchivoExcel, dsArchivoExcel);

                #endregion


                //CATALOGO AREA
                List myListCatalogoArea = myWeb.Lists.GetByTitle(catalogoArea);

                ListItemCollection listCatalogoArea = myListCatalogoArea.GetItems(CamlQuery.CreateAllItemsQuery());
                clientContext.Load(listCatalogoArea);
                clientContext.ExecuteQuery();

                //CATALOGO DEPARTAMENTE
                List myListCatalogoDepartamento = myWeb.Lists.GetByTitle(catalogoDepartamento);

                ListItemCollection listCatalogoDepartamento = myListCatalogoDepartamento.GetItems(CamlQuery.CreateAllItemsQuery());
                clientContext.Load(listCatalogoDepartamento);
                clientContext.ExecuteQuery();

                //CATALOGO DOC TYPE
                List myListCatalogoDocType = myWeb.Lists.GetByTitle(catalogoDocType);

                ListItemCollection listCatalogoDocType = myListCatalogoDocType.GetItems(CamlQuery.CreateAllItemsQuery());
                clientContext.Load(listCatalogoDocType);
                clientContext.ExecuteQuery();

                //CATALOGO SBU
                List myListCatalogoSBU = myWeb.Lists.GetByTitle(catalogoSBU);

                ListItemCollection listCatalogoSBU = myListCatalogoSBU.GetItems(CamlQuery.CreateAllItemsQuery());
                clientContext.Load(listCatalogoSBU);
                clientContext.ExecuteQuery();

                //CATALOGO CLIENTE
                List myListCatalogoCliente = myWeb.Lists.GetByTitle(catalogoCliente);

                ListItemCollection listCatalogoCliente = myListCatalogoCliente.GetItems(CamlQuery.CreateAllItemsQuery());
                clientContext.Load(listCatalogoCliente);
                clientContext.ExecuteQuery();

                #region ACTUALIZAR CODIGO ARCHIVO

                long maxId = listItems.Max(n => Convert.ToInt64(((ListItem)n).FieldValues["ID"]));
                long minId = listItems.Min(n => Convert.ToInt64(((ListItem)n).FieldValues["ID"]));

                var uno = listItems.Where(n => ((ListItem)n).FieldValues["DepartmentCode"] != null).ToList();


                string codigoDocumento = string.Empty;
                listItems.ToList().ForEach(n =>
                {
                    var area = ((ListItem)n).FieldValues["Area"];
                    var departamento = ((ListItem)n).FieldValues["Department"];
                    var tipoDocumento = ((ListItem)n).FieldValues["DocType"];
                    var id = ((ListItem)n).FieldValues["ID"];
                    long idc = 0;

                    idc = Convert.ToInt64(id) - 6525;
                    n["IDC"] = idc;

                    //n.Update();

                    if (area != null && departamento != null && tipoDocumento != null)
                    {

                        //clientContext.ExecuteQuery();


                        if (Convert.ToInt64(id) >= 1)
                        {
                            

                            codigoDocumento = ((FieldLookupValue)area).LookupValue.ToUpper().Substring(0, 3) + "-" +
                                            ((FieldLookupValue)departamento).LookupValue.ToUpper().Substring(0, 3) + "-" +
                                            ((FieldLookupValue)tipoDocumento).LookupValue.ToUpper().Substring(0, 3) + "-" +
                                            idc.ToString("0000");

                            n["DepartmentCode"] = codigoDocumento;

                            //n.Update();

                        }
                    }

                    n.Update();

                    if (Convert.ToInt64(id) % 100 == 0)
                    {
                        //n.Update();

                        clientContext.ExecuteQuery();
                    }

                    if (listItems.Count == Convert.ToInt64(id))
                    {
                        clientContext.ExecuteQuery();
                    }

                    
                });

                clientContext.ExecuteQuery();


                #endregion

                #region ACTUALIZAR AREA

                //ListItem AreaTot = listCatalogoArea.ToArray().Where(n => ((ListItem)n).FieldValues["Title"] != null && ((ListItem)n).FieldValues["Title"].ToString().ToUpper().Trim().Contains("SATELITES")).ToList().FirstOrDefault();

                //var res = listItems.Where(n => n.FieldValues["Area"] != null && ((FieldLookupValue)n.FieldValues["Area"]).LookupValue.ToUpper().Trim().Contains(AreaTot["Title"].ToString().ToUpper()));


                //foreach (ListItem item in res)
                //{
                //    ListItem Area = listCatalogoArea.ToArray().Where(n => ((ListItem)n).FieldValues["Title"] != null && ((ListItem)n).FieldValues["Title"].ToString().ToUpper().Trim().Contains("LOCALIDADES PPG")).ToList().FirstOrDefault();

                //    if (Area != null)
                //    {
                //        item["Area"] = Area;
                //        item.Update();
                //    }
                //    else
                //    {
                //        //Area = listCatalogoArea.ToList().FirstOrDefault();

                //        //item["Area"] = Area;

                //        continue;
                //    }
                //}

                //clientContext.ExecuteQuery();

                #endregion


                #region ACTUALIZAR DEPARTAMENTO

                //string departamento = "Powder";
                //string departamentoNuevo = "Pintura en Polvo";

                //ListItem DepartamentoTot = listCatalogoDepartamento.ToArray().Where(n => ((ListItem)n).FieldValues["Title"] != null && ((ListItem)n).FieldValues["Title"].ToString().ToUpper().Trim().Contains(departamento.ToUpper())).ToList().FirstOrDefault();

                //var resDep = listItems.Where(n => n.FieldValues["Department"] != null && ((FieldLookupValue)n.FieldValues["Department"]).LookupValue.ToUpper().Trim().Contains(DepartamentoTot["Title"].ToString().ToUpper()));


                //foreach (ListItem item in resDep)
                //{
                //    ListItem Departamento = listCatalogoDepartamento.ToArray().Where(n => ((ListItem)n).FieldValues["Title"] != null && ((ListItem)n).FieldValues["Title"].ToString().ToUpper().Trim().Contains(departamentoNuevo.ToUpper())).ToList().FirstOrDefault();

                //    if (Departamento != null)
                //    {
                //        item["Department"] = Departamento;
                //        item.Update();
                //    }
                //    else
                //    {
                //        //Area = listCatalogoArea.ToList().FirstOrDefault();

                //        //item["Area"] = Area;

                //        continue;
                //    }
                //}

                //clientContext.ExecuteQuery();

                #endregion

                //CARGAR METADATOS A SHARE POINT
                foreach (Archivo am in ListArchivo.ToList())//.Where(n => !string.IsNullOrEmpty(n.SBU) || !string.IsNullOrEmpty(n.DescripcionDocumento)))
                //foreach (Archivo am in ListArchivoEncontrado.ToList())
                {
                    try
                    {
                        ListItem item = listItems.ToArray().Where(n => ((ListItem)n).FieldValues["FileLeafRef"] != null && ((ListItem)n).FieldValues["FileLeafRef"].ToString() == am.Codigo).ToList().FirstOrDefault();

                        if (item != null)
                        {
                            try
                            {
                                //item.Update();
                                item.File.CheckOut();
                                clientContext.ExecuteQuery();
                            }
                            catch(Exception ex)
                            {

                            }

                            ////item.File.UndoCheckOut();

                            //item["Title"] = am.Codigo;

                            //item["Loop"] = "Si";
                            ////item["SBU"] = string.IsNullOrEmpty(am.SBU) ? "": am.SBU;

                            ////item["DepartmentCode"] = am.DepartamentoCodigo;
                            //////item["Cliente"] = string.IsNullOrEmpty(am.DescripcionDocumento) ? "": am.DescripcionDocumento;

                            //item["Revision"] = am.NumeroRevision;
                            ////item["Area"] = "";
                            ////item["Department"] = "";
                            ////item["DocType"] = "";

                            #region Catalogo Area

                            if (!string.IsNullOrEmpty(am.Area))
                            {
                                //ListItem Area = listCatalogoArea.ToArray().Where(n => ((ListItem)n).FieldValues["Title"] != null && ((ListItem)n).FieldValues["Title"].ToString().ToUpper().Trim().Contains(am.Area.ToUpper().Trim())).ToList().FirstOrDefault();

                                ListItem Area = listCatalogoArea.ToArray().Where(n => ((ListItem)n).FieldValues["Title"] != null && ((ListItem)n).FieldValues["Title"].ToString().ToUpper().Trim().Contains(am.Area.ToUpper().Trim())).ToList().FirstOrDefault();

                                if (Area != null)
                                {
                                    item["Area"] = Area;
                                }
                                else
                                {
                                    //Area = listCatalogoArea.ToList().FirstOrDefault();

                                    //item["Area"] = Area;

                                    continue;
                                }
                            }
                            else
                            {
                                //ListItem Area = listCatalogoArea.ToList().FirstOrDefault();

                                //item["Area"] = Area;

                                continue;
                            }

                            #endregion

                            //#region Catalogo Departamento

                            //if (!string.IsNullOrEmpty(am.Departamento))
                            //{
                            //    ListItem Departamento = listCatalogoDepartamento.ToArray().Where(n => ((ListItem)n).FieldValues["Title"] != null && ((ListItem)n).FieldValues["Title"].ToString().ToUpper().Trim().Contains(am.Departamento.ToUpper().Trim())).ToList().FirstOrDefault();

                            //    if (Departamento != null)
                            //    {
                            //        item["Department"] = Departamento;
                            //    }
                            //    else
                            //    {
                            //        //Departamento = listCatalogoDepartamento.ToList().FirstOrDefault();

                            //        //item["Department"] = Departamento;

                            //        continue;
                            //    }
                            //}
                            //else
                            //{
                            //    //ListItem Departamento = listCatalogoDepartamento.ToList().FirstOrDefault();

                            //    //item["Department"] = Departamento;

                            //    continue;
                            //}

                            //#endregion

                            //#region Catalogo TipoDocumento

                            //if (!string.IsNullOrEmpty(am.TipoDocumento))
                            //{
                            //    ListItem DocType = listCatalogoDocType.ToArray().Where(n => ((ListItem)n).FieldValues["Title"] != null && ((ListItem)n).FieldValues["Title"].ToString().ToUpper().Trim().Contains(am.TipoDocumento.ToUpper().Trim())).ToList().FirstOrDefault();

                            //    if (DocType != null)
                            //    {
                            //        item["DocType"] = DocType;
                            //    }
                            //    else
                            //    {
                            //        //DocType = listCatalogoDocType.ToList().FirstOrDefault();

                            //        //item["DocType"] = DocType;

                            //        //continue;

                            //        item["DocType"] = null;
                            //    }
                            //}
                            //else
                            //{
                            //    //ListItem DocType = listCatalogoDocType.ToList().FirstOrDefault();

                            //    //item["DocType"] = DocType;

                            //    item["DocType"] = null;
                            //}

                            //#endregion

                            //#region SBU

                            //if (!string.IsNullOrEmpty(am.SBU))
                            //{
                            //    ListItem SBU = listCatalogoSBU.ToArray().Where(n => ((ListItem)n).FieldValues["Title"] != null && ((ListItem)n).FieldValues["Title"].ToString().ToUpper().Trim().Contains(am.SBU.ToUpper().Trim())).ToList().FirstOrDefault();

                            //    if (SBU != null)
                            //    {
                            //        item["SBU"] = SBU;
                            //    }
                            //    else
                            //    {
                            //        item["SBU"] = null;
                            //    }
                            //}
                            //else
                            //{
                            //    item["SBU"] = null;
                            //}

                            //#endregion

                            //#region Catalogo Cliente

                            //if (!string.IsNullOrEmpty(am.Cliente))
                            //{
                            //    ListItem Clientes = listCatalogoCliente.ToArray().Where(n => ((ListItem)n).FieldValues["Title"] != null && ((ListItem)n).FieldValues["Title"].ToString().ToUpper().Trim().Contains(am.Cliente.ToUpper().Trim())).ToList().FirstOrDefault();

                            //    if (Clientes != null)
                            //    {
                            //        item["Cliente"] = Clientes;
                            //    }
                            //    else
                            //    {
                            //        item["Cliente"] = null;
                            //    }
                            //}
                            //else
                            //{
                            //    item["Cliente"] = null;
                            //}

                            //#endregion

                            //ACTUALIZAR INFORMACION DE METADATOS
                            //item.File.UndoCheckOut();
                            item.Update();
                            //clientContext.ExecuteQuery();

                            //REALIZAR CHECKOUT PARA TOMAR LOS ARCHIVOS
                            //item.File.CheckOut();
                            //clientContext.ExecuteQuery();

                            //REALIZAR CHECKIN DE LOS ARCHIVOS
                            item.File.CheckIn("", CheckinType.OverwriteCheckIn);
                            clientContext.ExecuteQuery();

                            dtArchivoExcel.Rows.Add(string.IsNullOrEmpty(am.Area) ? " ": am.Area,
                                                    string.IsNullOrEmpty(am.Departamento) ? " ": am.Departamento,
                                                    string.IsNullOrEmpty(am.TipoDocumento) ? " ": am.TipoDocumento, 
                                                    string.IsNullOrEmpty(am.DepartamentoCodigo) ? " " : am.DepartamentoCodigo,
                                                    string.IsNullOrEmpty(am.Codigo) ? " ": am.Codigo,
                                                    string.IsNullOrEmpty(am.NombreDocumento) ? " ": am.NombreDocumento,
                                                    string.IsNullOrEmpty(am.DescripcionDocumento) ? " " : am.DescripcionDocumento,
                                                    string.IsNullOrEmpty(am.NumeroRevision) ? " ": am.NumeroRevision,
                                                    string.IsNullOrEmpty(am.FCambioFijo) ? " ": am.FCambioFijo,
                                                    string.IsNullOrEmpty(am.FCambioFrecuente) ? " ": am.FCambioFrecuente,
                                                    string.IsNullOrEmpty(am.SBU) ? " " : am.SBU,
                                                    string.IsNullOrEmpty(am.Cliente) ? " " : am.Cliente);

                            //if(dtArchivoExcel.Rows.Count % 200 == 0)
                            //{
                                //clientContext.ExecuteQuery();
                            //}
                        }
                    }
                    catch(Exception ex)
                    {
                        
                    }
                }

                //clientContext.ExecuteQuery();


                //CREAR ARCHIVO EXCEL

                dsArchivoExcel.Tables.Add(dtArchivoExcel);

                ExcelLibrary.DataSetHelper.CreateWorkbook(pathArchivoExcel, dsArchivoExcel);

                ////EXPORTAR CATALOGOS
                ////AREA
                //dtArea.Columns.Add("Title");
                //dtArea.Columns.Add("Code");
                //foreach(ListItem item in listCatalogoArea)
                //{
                //    dtArea.Rows.Add(item["Title"], item["Code"]);
                //}

                ////DEPARTAMENTO
                //dtDepartamento.Columns.Add("Title");
                //dtDepartamento.Columns.Add("Area");
                //dtDepartamento.Columns.Add("Code");
                //foreach (ListItem item in listCatalogoDepartamento)
                //{
                //    dtDepartamento.Rows.Add(item["Title"], ((FieldLookupValue)item["Area"]).LookupValue, item["b8ph"]);
                //}
                ////DOC TYPE
                //dtDocType.Columns.Add("Title");
                //foreach (ListItem item in listCatalogoDocType)
                //{
                //    dtDocType.Rows.Add(item["Title"]);
                //}

                //dsCatalogos.Tables.Add(dtArea);
                //dsCatalogos.Tables.Add(dtDepartamento);
                //dsCatalogos.Tables.Add(dtDocType);

                //ExcelLibrary.DataSetHelper.CreateWorkbook(pathArchivoExcel, dsCatalogos);
            }
            catch(Exception ex)
            {
                Console.WriteLine("Mensaje: " + ex.Message + ", Source: " + ex.Source + ", StackTrace: " + ex.StackTrace);
            }
        }
    }
}
