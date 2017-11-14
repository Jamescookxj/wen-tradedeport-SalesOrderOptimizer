using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.Odbc;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using mshtml;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace WindowsFormsApp1
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        //public string bigcitylist = "Auckland,Wellington,Christchurch,Hamilton,Tauranga,Napier-Hastings,Dunedin,Palmerston North,Nelson,Rotorua,New Plymouth,Whangarei,Invercargill,Whanganui,Gisborne,".ToUpper();
        public string citylist = "Ahaura,Ahipara,Ahititi,Ahuroa,Akaroa,Akitio,Albany,Albert Town,Albury,Alexandra,Allanton,Amberley,Anakiwa,Aranga,Aramoana,Arapohue,Arrowtown,Arundel,Ashburton,Ashhurst,Ashley,Auckland,Auroa,Awanui,Bedit,Balclutha,Balfour,Barrhill,Barrytown,Beachlands,Beaumont,Bell Block,Benhar,Benneydale,Bideford,Blackball,Blenheim,Bluff,Brighton,Brightwater,Broadwood,Bulls,Bunnythorpe,Burnt Hill,Cedit,Cambridge,Canvastown,Carterton,Cheviot,Christchurch,Clarkville,Clarksville,Clevedon,Clinton,Clive,Clyde,Coatesville,Collingwood,Colville,Coopers Creek,Coroglen,Coromandel,Cromwell,Culverden,Cust,Dedit,Dairy Flat,Dannevirke,Darfield,Dargaville,Dipton,Dobson,Drury,Dunedin,Duntroon,Eedit,Eastbourne,Edendale,Edgecumbe,Egmont Village,Eketahuna,Eltham,Ettrick,Eyrewell Forest,Fedit,Fairhall,Fairlie,Featherston,Feilding,Fernside,Flaxmere,Flaxton,Fox Glacier,Foxton,Foxton Beach,Frankton, Otago,Frankton, Waikato,Franz Josef,Gedit,Geraldine,Gisborne,Glenorchy,Glentui,Gore,Granity,Greymouth,Greytown,Grovetown,Gummies Bush,Hedit,Haast,Hakataramea,Halcombe,Hamilton,Hampden,Hanmer Springs,Hari Hari,Hastings,Haumoana,Haupiri,Havelock,Havelock North,Hāwea,Hawera,Helensville,Henley,Herbert,Herekino,Hikuai,Hikurangi,Hikutaia,Hinuera,Hokitika,Hope,Horeke,Houhora,Howick,Huapai,Huiakama,Huirangi,Hukerenui,Hunterville,Huntly,Hurleyville,Iedit,Inangahua Junction,Inglewood,Invercargill,Jedit,Jack's Point,Jacobs River,Kedit,Kaeo,Kaiapoi,Kaihu,Kaikohe,Kaikoura,Kaimata,Kaingaroa,Kakanui,Kaipara Flats,Kairaki,Kaitaia,Kaitangata,Kaiwaka,Kakaramea,Kaniere,Kaponga,Karamea,Karetu,Karitane,Katikati,Kaukapakapa,Kauri,Kawakawa,Kawerau,Kennedy Bay,Kerikeri,Kihikihi,Kinloch,Kingston,Kirwee,Kohukohu,Koitiata,Kokatahi,Kokopu,Koromiko,Kumara,Kumeu,Kurow,Ledit,Lauriston,Lawrence,Leeston,Leigh,Lepperton,Levin,Lincoln,Linkwater,Little River,Loburn,Lower Hutt,Luggate,Lumsden,Lyttelton,Medit,Makahu,Manaia, Taranaki,Manaia, Coromandel,Manakau,Manapouri,Mangakino,Mangamuka,Mangatoki,Mangawhai,Manukau,Manurewa,Manutahi,Mapua,Maraetai,Marco,Maromaku,Marsden Bay,Martinborough,Marton,Maruia,Masterton,Matakana,Matakohe,Matamata,Matapu,Matarangi,Matarau,Matata,Mataura,Matihetihe,Maungakaramea,Maungatapere,Maungaturoto,Mayfield,Meremere,Methven,Middlemarch,Midhirst,Millers Flat,Milton,Mimi,Minginui,Moana,Moawhango,Moenui,Moeraki,Moerewa,Mokau,Mokoia,Morrinsville,Mosgiel,Mossburn,Motatau,Motueka,Mount Maunganui,Mount Somers,Murchison,Murupara,Nedit,Napier,Naseby,Nelson,New Brighton,New Plymouth,Normanby,Ngaere,Ngamatapouri,Ngapara,Ngaruawahia,Ngataki,Ngatea,Ngongotaha,Ngunguru,Nightcaps,Norfolk,Norsewood,Oedit,Oakura,Oamaru,Oban,Ohakune,Ohaeawai,Ohangai,Ohoka,Ōhope Beach,Ohura,Okaihau,Okato,Okuku,Omanaia,Omarama,Omata,Omokoroa,Onewhero,Opononi,Opotiki,Opua,Opunake,Oratia,Orewa,Oromahoe,Oruaiti,Otaika,Otaki,Otakou,Otautau,Otiria,Otorohanga,Owaka,Oxford,Pedit,Paekakariki,Paeroa,Pahiatua,Paihia,Pakaraka,Pakiri,Pakotai,Palmerston,Palmerston North,Pamapuria,Panguru,Papakura,Papamoa,Paparoa,Paparore,Papatoetoe,Parakai,Paraparaumu,Paremoremo,Pareora,Paroa,Parua Bay,Patea,Pauanui,Pauatahanui,Pegasus,Peka Peka,Pembroke,Peria,Petone,Picton,Piopio,Pipiwai,Pirongia,Pleasant Point,Plimmerton,Pokeno,Porirua,Portland,Poroti,Port Chalmers,Portobello,Pukekohe,Pukerua Bay,Pukeuri,Pukepoto,Punakaiki,Purua,Putaruru,Putorino,Qedit,Queenstown,Redit,Raetihi,Raglan,Rahotu,Rai Valley,Rakaia,Ramarama,Ranfurly,Rangiora,Rapaura,Ratapiko,Raumati,Rawene,Rawhitiroa,Reefton,Renwick,Reporoa,Richmond,Riverhead,Riverlands,Riversdale,Riversdale Beach,Riverton,Riwaka,Rolleston,Ross,Rotorua,Roxburgh,Ruatoria,Ruakaka,Ruawai,Runanga,Russell,Sedit,Saint Andrews,Saint Arnaud,Saint Bathans,Sanson,Seacliff,Seddon,Seddonville,Sefton,Sheffield,Shannon,Silverdale,Snells Beach,Springfield,Springston,Spring Creek,Stirling,Stratford,Swannanoa,Tedit,Taharoa,Taieri Mouth,Taihape,Taipa-Mangonui,Tairua,Takaka,Tangiteroria,Tapanui,Tapu,Tangowahine,Tapawera,Tapora,Taradale,Tauhoa,Taumarunui,Taupaki,Taupo,Tauranga,Tauraroa,Tautoro,Te Anau,Te Arai,Te Aroha,Te Awamutu,Te Awanga,Te Hapua,Te Horo,Te Kao,Te Kauwhata,Te Kopuru,Te Kuiti,Te Poi,Te Puke,Te Puru,Temuka,Te Rerenga,Thames,Tikorangi,Timaru,Tinopai,Tinwald,Tirau,Titoki,Tokarahi,Toko,Tokanui,Tokoroa,Tolaga Bay,Tomarata,Towai,Tuahiwi,Tuai,Tuakau,Tuamarina,Tuatapere,Turangi,Twizel,Uedit,Umawera,Upper Hutt,Upper Moutere,Urenui,Uruti,Vedit,View Hill,Wedit,Waddington,Waiheke Island,Waipango,Waharoa,Waiharara,Waihi,Waihi Beach,Waihola,Waikaia,Waikaka,Waikanae,Waikawa, Marlborough,Waikawa, Southland,Waikouaiti,Waikuku,Waikuku Beach,Waima,Waimangaroa,Waimate,Waimate North,Waimauku,Wainui,Wainuiomata,Waioneke,Waiouru,Waiotira,Waipawa,Waipukurau,Wairakei,Wairau Valley,Wairoa,Waitahuna,Waitara,Waitaria Bay,Waitati,Waitoa,Waitoki,Waitoriki,Waitotara,Waiuku,Waiwera,Wakefield,Wallacetown,Walton,Waverley,Wanaka,Ward,Wardville,Warrington,Warkworth,Wellington,Wellsford,Westport,Weston,Whakatane,Whakamaru,Whananaki,Whangamata,Whangamomona,Whanganui,Whangarei,Whangarei Heads,Whangaruru,Whataroa,Whatuwhiwhi,Whenuakite,Whenuakura,Whiritoa,Whitford,Whitby,Whitianga,Willowby,Wimbledon,Winchester,Windsor,Windwhistle,Winscombe,Winton,Woodend,Woodend Beach,Woodhill,Woodville,Wyndham,".ToUpper();
        private void button1_Click(object sender, EventArgs e)
        {//"SELECT  ListId  ,0 as weight, 0 as volume from Employee"
            OdbcConnection con = new OdbcConnection("DSN=QuickBooks Data QRemote");
            con.Open();
            string TxnDate = ""; string years = ""; string months = ""; string days = "";
            years = dateTimePicker1.Value.Year.ToString(); months = dateTimePicker1.Value.Month.ToString(); days = dateTimePicker1.Value.Day.ToString();
            if (days.Length==1) { days = '0' + days; }
            TxnDate = years+'-'+ months+'-'+ days;
            string sql = "Select TxnDate, RefNumber, ShipAddressBlockAddr1, ShipAddressBlockAddr2, ShipAddressBlockAddr3, ShipAddressBlockAddr4, ShipAddressBlockAddr5, Memo from SalesOrder where RefNumber  = '" + textBox1.Text.Trim() + "'";
            OdbcDataAdapter dAdapter = new OdbcDataAdapter(sql , con);
            DataTable result = new DataTable();
           // result.Columns.Add("check", typeof(bool));

            dAdapter.Fill(result);
            for (int i = 0; i < result.Rows.Count; i++)
            {
                textBox2.Text = result.Rows[i]["TxnDate"].ToString();
                textBox3.Text = result.Rows[i]["RefNumber"].ToString();
                textBox6.Text = result.Rows[i]["ShipAddressBlockAddr1"].ToString();
                textBox5.Text = result.Rows[i]["ShipAddressBlockAddr2"].ToString();
                textBox4.Text = result.Rows[i]["ShipAddressBlockAddr3"].ToString();
                textBox13.Text = result.Rows[i]["ShipAddressBlockAddr4"].ToString();
                textBox12.Text = result.Rows[i]["ShipAddressBlockAddr5"].ToString();
                textBox8.Text = result.Rows[i]["memo"].ToString();
            }
            //this.dataGridView1.AutoGenerateColumns = false;
            // this.dataGridView1.DataSource = result;

            con.Close();
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void webBrowser1_DocumentCompleted(object sender, WebBrowserDocumentCompletedEventArgs e)
        {
            if (webBrowser1.Url.ToString()== "http://webnotes.streamlinefreight.co.nz/") {
                HtmlElement headElement = webBrowser1.Document.GetElementsByTagName("head")[0];
                HtmlElement scriptElement = webBrowser1.Document.CreateElement("script");
                IHTMLScriptElement element = (IHTMLScriptElement)scriptElement.DomElement;
                element.text = "function sayHello() { " +
                    "document.getElementById('mainContent_boxCustCode').value = '13731';  document.getElementById('mainContent_boxLogin').value = '13731'; document.getElementById('mainContent_boxPassword').value = '13731';   document.getElementById('mainContent_Button1').click();}";
                headElement.AppendChild(scriptElement);
                webBrowser1.Document.InvokeScript("sayHello");
            }
            if (webBrowser1.Url.ToString() == "http://webnotes.streamlinefreight.co.nz/ConsignmentManagement/Consignment.aspx")
            {
                if (dataGridView1.CurrentRow != null)
                {
                    string city = ""; string city2 = "";
                    if (dataGridView1.CurrentRow.Cells.Count > 0)
                    {
                        string ordernumber = dataGridView1.CurrentRow.Cells[2].Value.ToString();

                        string str1 = dataGridView1.CurrentRow.Cells[3].Value.ToString(); if (citylist.IndexOf((str1.ToUpper() + ",")) >= 0 && str1 != "") city = str1;
                        if (str1.IndexOf(',') > 0) city2 = str1.Substring(0, str1.IndexOf(',')).ToUpper().Trim();
                        if (citylist.IndexOf((city2 + ",")) >= 0 && city2 != "") city2 = city2; else city2 = "";
                        if (city2 == "" && str1.IndexOf(' ') > 0) city2 = str1.Substring(0, str1.IndexOf(' ')).ToUpper().Trim();
                        if (citylist.IndexOf((city2 + ",")) >= 0 && city2 != "") city2 = city2; else city2 = "";

                        string str2 = dataGridView1.CurrentRow.Cells[4].Value.ToString(); if (citylist.IndexOf((str2.ToUpper() + ",")) >= 0 && str2 != "") city = str2;
                        if (str2.IndexOf(',') > 0) city2 = str2.Substring(0, str2.IndexOf(',')).ToUpper().Trim();
                        if (citylist.IndexOf((city2 + ",")) >= 0 && city2 != "") city2 = city2; else city2 = "";
                        if (city2 == "" && str2.IndexOf(' ') > 0) city2 = str2.Substring(0, str2.IndexOf(' ')).ToUpper().Trim();
                        if (citylist.IndexOf((city2 + ",")) >= 0 && city2 != "") city2 = city2; else city2 = "";

                        string str3 = dataGridView1.CurrentRow.Cells[5].Value.ToString(); if (citylist.IndexOf((str3.ToUpper() + ",")) >= 0 && str3 != "") city = str3;
                        if (str3.IndexOf(',') > 0) city2 = str3.Substring(0, str3.IndexOf(',')).ToUpper().Trim();
                        if (citylist.IndexOf((city2 + ",")) >= 0 && city2 != "") city2 = city2; else city2 = "";
                        if (city2 == "" && str3.IndexOf(' ') > 0) city2 = str3.Substring(0, str3.IndexOf(' ')).ToUpper().Trim();
                        if (citylist.IndexOf((city2 + ",")) >= 0 && city2 != "") city2 = city2; else city2 = "";

                        string str4 = dataGridView1.CurrentRow.Cells[6].Value.ToString(); if (citylist.IndexOf((str4.ToUpper() + ",")) >= 0 && str4 != "") city = str4;
                        if (str4.IndexOf(',') > 0) city2 = str4.Substring(0, str4.IndexOf(',')).ToUpper().Trim();
                        if (citylist.IndexOf((city2 + ",")) >= 0 && city2 != "") city2 = city2; else city2 = "";
                        if (city2 == "" && str4.IndexOf(' ') > 0) city2 = str4.Substring(0, str4.IndexOf(' ')).ToUpper().Trim();
                        if (citylist.IndexOf((city2 + ",")) >= 0 && city2 != "") city2 = city2; else city2 = "";

                        string str5 = dataGridView1.CurrentRow.Cells[7].Value.ToString(); if (citylist.IndexOf((str5.ToUpper() + ",")) >= 0 && str5 != "") city = str5;
                        if (str5.IndexOf(',') > 0) city2 = str5.Substring(0, str5.IndexOf(',')).ToUpper().Trim();
                        if (citylist.IndexOf((city2 + ",")) >= 0 && city2 != "") city2 = city2; else city2 = "";
                        if (city2 == "" && str5.IndexOf(' ') > 0) city2 = str5.Substring(0, str5.IndexOf(' ')).ToUpper().Trim();
                        if (citylist.IndexOf((city2 + ",")) >= 0 && city2 != "") city2 = city2; else city2 = "";

                        if (city == "" && city2 != "") city = city2;
                        if (city == "") city = str5;

                         
                        HtmlElement orderRef1 = webBrowser1.Document.GetElementById("mainContent_txtOrderNo");
                        if (orderRef1 != null) orderRef1.SetAttribute("value", "SO " + ordernumber);
                        HtmlElement addressL1 = webBrowser1.Document.GetElementById("mainContent_txtReceiverName");
                        if (addressL1 != null) addressL1.SetAttribute("value", str1+" "+ str2);
                        HtmlElement addressL2 = webBrowser1.Document.GetElementById("mainContent_txtReceiverAddr1");
                        if (addressL2 != null) addressL2.SetAttribute("value", str3);
                        HtmlElement addressL3 = webBrowser1.Document.GetElementById("mainContent_txtReceiverAddr2");
                        if (addressL3 != null) addressL3.SetAttribute("value", str4);                         
                        HtmlElement addressL6_Location = webBrowser1.Document.GetElementById("mainContent_loc_receiver");
                        if (addressL6_Location != null) addressL6_Location.SetAttribute("value", city);
                           


                    }
                }
        }
                
            if (webBrowser1.Url.ToString() == "https://my.castleparcels.co.nz/Account/Login" || webBrowser1.Url.ToString() == "https://my.castleparcels.co.nz/Account/Login?ReturnUrl=%2fCourier%2fSend%2feDespatchIT")
            {
                HtmlElement headElement = webBrowser1.Document.GetElementsByTagName("head")[0];
                HtmlElement scriptElement = webBrowser1.Document.CreateElement("script");
                IHTMLScriptElement element = (IHTMLScriptElement)scriptElement.DomElement;
                element.text = "function sayHello() { " +
                    " document.getElementById('UserName').value = 'warehouse306@tradedepot.co.nz'; document.getElementById('Password').value = 'tradedepot';   }";
                headElement.AppendChild(scriptElement);
                webBrowser1.Document.InvokeScript("sayHello");
                webBrowser1.Document.Forms[0].InvokeMember("submit");
            }
            if (webBrowser1.Url.ToString() == "https://my.castleparcels.co.nz/Courier/Send/eDespatchIT") {
                
                    string city = ""; string city2 = ""; string str1 = ""; string str2 = ""; string str3 = ""; string str4 = ""; string str5 = ""; string ordernumber = ""; string quantity = ""; string weight = ""; string volume = "";
                ordernumber = textBox3.Text.Trim();
                    str1 = textBox6.Text.Trim(); if (citylist.IndexOf((str1.ToUpper() + ",")) >= 0 && str1 != "") city = str1;
                    if (str1.IndexOf(',') > 0) city2 = str1.Substring(0, str1.IndexOf(',')).ToUpper().Trim();
                    if (citylist.IndexOf((city2 + ",")) >= 0 && city2 != "") city2 = city2; else city2 = "";
                    if (city2 == "" && str1.IndexOf(' ') > 0) city2 = str1.Substring(0, str1.IndexOf(' ')).ToUpper().Trim();
                    if (citylist.IndexOf((city2 + ",")) >= 0 && city2 != "") city2 = city2; else city2 = "";

                    str2 = textBox5.Text.Trim(); if (citylist.IndexOf((str2.ToUpper() + ",")) >= 0 && str2 != "") city = str2;
                    if (str2.IndexOf(',') > 0) city2 = str2.Substring(0, str2.IndexOf(',')).ToUpper().Trim();
                    if (citylist.IndexOf((city2 + ",")) >= 0 && city2 != "") city2 = city2; else city2 = "";
                    if (city2 == "" && str2.IndexOf(' ') > 0) city2 = str2.Substring(0, str2.IndexOf(' ')).ToUpper().Trim();
                    if (citylist.IndexOf((city2 + ",")) >= 0 && city2 != "") city2 = city2; else city2 = "";

                    str3 = textBox4.Text.Trim(); if (citylist.IndexOf((str3.ToUpper() + ",")) >= 0 && str3 != "") city = str3;
                    if (str3.IndexOf(',') > 0) city2 = str3.Substring(0, str3.IndexOf(',')).ToUpper().Trim();
                    if (citylist.IndexOf((city2 + ",")) >= 0 && city2 != "") city2 = city2; else city2 = "";
                    if (city2 == "" && str3.IndexOf(' ') > 0) city2 = str3.Substring(0, str3.IndexOf(' ')).ToUpper().Trim();
                    if (citylist.IndexOf((city2 + ",")) >= 0 && city2 != "") city2 = city2; else city2 = "";

                    str4 = textBox13.Text.Trim(); if (citylist.IndexOf((str4.ToUpper() + ",")) >= 0 && str4 != "") city = str4;
                    if (str4.IndexOf(',') > 0) city2 = str4.Substring(0, str4.IndexOf(',')).ToUpper().Trim();
                    if (citylist.IndexOf((city2 + ",")) >= 0 && city2 != "") city2 = city2; else city2 = "";
                    if (city2 == "" && str4.IndexOf(' ') > 0) city2 = str4.Substring(0, str4.IndexOf(' ')).ToUpper().Trim();
                    if (citylist.IndexOf((city2 + ",")) >= 0 && city2 != "") city2 = city2; else city2 = "";

                    str5 = textBox12.Text.Trim(); if (citylist.IndexOf((str5.ToUpper() + ",")) >= 0 && str5 != "") city = str5;
                    if (str5.IndexOf(',') > 0) city2 = str5.Substring(0, str5.IndexOf(',')).ToUpper().Trim();
                    if (citylist.IndexOf((city2 + ",")) >= 0 && city2 != "") city2 = city2; else city2 = "";
                    if (city2 == "" && str5.IndexOf(' ') > 0) city2 = str5.Substring(0, str5.IndexOf(' ')).ToUpper().Trim();
                    if (citylist.IndexOf((city2 + ",")) >= 0 && city2 != "") city2 = city2; else city2 = "";

                    if (city == "" && city2 != "") city = city2;
                    if (city == "") city = str5;
                    quantity = textBox9.Text; weight = textBox7.Text; volume = textBox10.Text;

                    HtmlElement orderRef1 = webBrowser1.Document.GetElementById("orderRef1");
                    if (orderRef1 != null) orderRef1.SetAttribute("value", "SO " + ordernumber);
                    HtmlElement addressL1 = webBrowser1.Document.GetElementById("addressL1");
                    if (addressL1 != null) addressL1.SetAttribute("value", str1);
                    HtmlElement addressL2 = webBrowser1.Document.GetElementById("addressL2");
                    if (addressL2 != null) addressL2.SetAttribute("value", str2);
                    HtmlElement addressL3 = webBrowser1.Document.GetElementById("addressL3");
                    if (addressL3 != null) addressL3.SetAttribute("value", str3);
                    HtmlElement addressL4 = webBrowser1.Document.GetElementById("addressL4");
                    if (addressL4 != null) addressL4.SetAttribute("value", str4);
                    HtmlElement addressL5 = webBrowser1.Document.GetElementById("addressL5");
                    if (addressL5 != null) addressL5.SetAttribute("value", str5);
                    HtmlElement addressL6_Location = webBrowser1.Document.GetElementById("addressL6_Location");
                    if (addressL6_Location != null) addressL6_Location.SetAttribute("value", city);
                    HtmlElement BasePlus_Quantity = webBrowser1.Document.GetElementById("BasePlus_Quantity");
                    if (BasePlus_Quantity != null) addressL6_Location.SetAttribute("value", quantity);
                    HtmlElement BasePlus_Amount = webBrowser1.Document.GetElementById("BasePlus_Amount");
                    if (BasePlus_Amount != null) addressL6_Location.SetAttribute("value", volume);
              //  HtmlElement BasePlus_Quantity = webBrowser1.Document.GetElementById("BasePlus_Quantity");
             //   if (BasePlus_Quantity != null) addressL6_Location.SetAttribute("value", quantity);
                /*  if (dataGridView1.CurrentRow != null)
                     { if (dataGridView1.CurrentRow.Cells.Count > 0)
                     {
                     string ordernumber = dataGridView1.CurrentRow.Cells[2].Value.ToString();
                     string str1 = dataGridView1.CurrentRow.Cells[3].Value.ToString(); if (citylist.IndexOf((str1.ToUpper() + ",")) >= 0 && str1 != "") city = str1;
                     if (str1.IndexOf(',') > 0) city2 = str1.Substring(0, str1.IndexOf(',')).ToUpper().Trim();
                     if (citylist.IndexOf((city2 + ",")) >= 0 && city2 != "") city2 = city2; else city2 = "";
                     if (city2 == "" && str1.IndexOf(' ') > 0) city2 = str1.Substring(0, str1.IndexOf(' ')).ToUpper().Trim();
                     if (citylist.IndexOf((city2 + ",")) >= 0 && city2 != "") city2 = city2; else city2 = "";

                     string str2 = dataGridView1.CurrentRow.Cells[4].Value.ToString(); if (citylist.IndexOf((str2.ToUpper() + ",")) >= 0 && str2 != "") city = str2;
                     if (str2.IndexOf(',') > 0) city2 = str2.Substring(0, str2.IndexOf(',')).ToUpper().Trim();
                     if (citylist.IndexOf((city2 + ",")) >= 0 && city2 != "") city2 = city2; else city2 = "";
                     if (city2 == "" && str2.IndexOf(' ') > 0) city2 = str2.Substring(0, str2.IndexOf(' ')).ToUpper().Trim();
                     if (citylist.IndexOf((city2 + ",")) >= 0 && city2 != "") city2 = city2; else city2 = "";

                     string str3 = dataGridView1.CurrentRow.Cells[5].Value.ToString(); if (citylist.IndexOf((str3.ToUpper() + ",")) >= 0 && str3 != "") city = str3;
                     if (str3.IndexOf(',') > 0) city2 = str3.Substring(0, str3.IndexOf(',')).ToUpper().Trim();
                     if (citylist.IndexOf((city2 + ",")) >= 0 && city2 != "") city2 = city2; else city2 = "";
                     if (city2 == "" && str3.IndexOf(' ') > 0) city2 = str3.Substring(0, str3.IndexOf(' ')).ToUpper().Trim();
                     if (citylist.IndexOf((city2 + ",")) >= 0 && city2 != "") city2 = city2; else city2 = "";

                     string str4 = dataGridView1.CurrentRow.Cells[6].Value.ToString(); if (citylist.IndexOf((str4.ToUpper() + ",")) >= 0 && str4 != "") city = str4;
                     if (str4.IndexOf(',') > 0) city2 = str4.Substring(0, str4.IndexOf(',')).ToUpper().Trim();
                     if (citylist.IndexOf((city2 + ",")) >= 0 && city2 != "") city2 = city2; else city2 = "";
                     if (city2 == "" && str4.IndexOf(' ') > 0) city2 = str4.Substring(0, str4.IndexOf(' ')).ToUpper().Trim();
                     if (citylist.IndexOf((city2 + ",")) >= 0 && city2 != "") city2 = city2; else city2 = "";

                     string str5 = dataGridView1.CurrentRow.Cells[7].Value.ToString(); if (citylist.IndexOf((str5.ToUpper() + ",")) >= 0 && str5 != "") city = str5;
                     if (str5.IndexOf(',') > 0) city2 = str5.Substring(0, str5.IndexOf(',')).ToUpper().Trim();
                     if (citylist.IndexOf((city2 + ",")) >= 0 && city2 != "") city2 = city2; else city2 = "";
                     if (city2 == "" && str5.IndexOf(' ') > 0) city2 = str5.Substring(0, str5.IndexOf(' ')).ToUpper().Trim();
                     if (citylist.IndexOf((city2 + ",")) >= 0 && city2 != "") city2 = city2; else city2 = "";

                     if (city == "" && city2 != "") city = city2;
                     if (city == "") city = str5;
                    }
                    }*/
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
           
        }

        private void button3_Click(object sender, EventArgs e)
        {
  
        }

        private void Form1_Load(object sender, EventArgs e)
        {
         //   webBrowser1.Navigate("https://my.castleparcels.co.nz/Account/Login");
            webBrowser1.Navigate("http://webnotes.streamlinefreight.co.nz/");
        }

        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            if (dataGridView1.CurrentRow!=null ) {
                if (dataGridView1.CurrentRow.Cells.Count > 0) {

                }
                   // MessageBox.Show(dataGridView1.CurrentRow.Cells[4].Value.ToString());
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {          
           // webBrowser1.Navigate("https://my.castleparcels.co.nz/Courier/Send/eDespatchIT");
        }

        private void button5_Click(object sender, EventArgs e)
        {
          //  webBrowser1.Navigate("http://webnotes.streamlinefreight.co.nz/");
        }

        private void textBox9_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar != '\b' && !Char.IsDigit(e.KeyChar) && e.KeyChar != '.')
            {
                e.Handled = true;
            }
        }

        private void textBox7_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar != '\b' && !Char.IsDigit(e.KeyChar) && e.KeyChar != '.')
            {
                e.Handled = true;
            }
        }

        private void textBox10_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar != '\b' && !Char.IsDigit(e.KeyChar) && e.KeyChar != '.')
            {
                e.Handled = true;
            }
        }

        private void button2_Click_1(object sender, EventArgs e)
        {
            OdbcConnection con = new OdbcConnection("DSN=QuickBooks Data QRemote");
            con.Open();
            string TxnDate = ""; string years = ""; string months = ""; string days = "";
            years = dateTimePicker1.Value.Year.ToString(); months = dateTimePicker1.Value.Month.ToString(); days = dateTimePicker1.Value.Day.ToString();
            if (days.Length == 1) { days = '0' + days; }
            TxnDate = years + '-' + months + '-' + days;
            string minute = "";
            //SELECT * FROM SalesOrder where TxnDate = {d '2017-11-06'} and TimeModified >= {ts '2017-11-06 13:00:00.000'}
            minute = (DateTime.Now.Minute - numericUpDown1.Value).ToString().Trim() ;
             if (minute.Length == 1) minute = "0" + minute;
            minute = DateTime.Now.Year.ToString() + "-" + DateTime.Now.Month.ToString().PadLeft(2,'0') + "-" + (DateTime.Now.Day-9).ToString().PadLeft(2, '0') + " " + DateTime.Now.Hour.ToString().PadLeft(2, '0') + ":" + minute + ":" + DateTime.Now.Second.ToString().PadLeft(2, '0') + ".000";
            //  string sql = "Select TimeCreated,TxnDate, RefNumber, ShipAddressBlockAddr1, ShipAddressBlockAddr2, ShipAddressBlockAddr3, ShipAddressBlockAddr4, ShipAddressBlockAddr5, Memo from SalesOrder  where TimeCreated> {ts '" + minute+ "'} and TxnNumber>335000 order by TimeCreated desc";
            string sql = "Select TimeCreated,TxnDate, RefNumber, ShipAddressBlockAddr1, ShipAddressBlockAddr2, ShipAddressBlockAddr3, ShipAddressBlockAddr4, ShipAddressBlockAddr5, Memo from SalesOrder  where TxnDate >= ({fn CURDATE()})    order by TimeCreated desc";
            OdbcDataAdapter dAdapter = new OdbcDataAdapter(sql, con);
            DataTable result = new DataTable();
            result.Columns.Add("check", typeof(bool));

            dAdapter.Fill(result);
            this.dataGridView1.AutoGenerateColumns = false;
            this.dataGridView1.DataSource = result;

            con.Close();
        }

        private void dataGridView1_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            textBox2.Text = dataGridView1.CurrentRow.Cells[1].Value.ToString();
            textBox3.Text = dataGridView1.CurrentRow.Cells[2].Value.ToString();
            textBox6.Text = dataGridView1.CurrentRow.Cells[3].Value.ToString();
            textBox5.Text = dataGridView1.CurrentRow.Cells[4].Value.ToString();
            textBox4.Text = dataGridView1.CurrentRow.Cells[5].Value.ToString();
            textBox13.Text = dataGridView1.CurrentRow.Cells[6].Value.ToString();
            textBox12.Text = dataGridView1.CurrentRow.Cells[7].Value.ToString();
            textBox8.Text = dataGridView1.CurrentRow.Cells[8].Value.ToString();
        }

        private void button4_Click_1(object sender, EventArgs e)
        {
            webBrowser1.Navigate("https://my.castleparcels.co.nz/Courier/Send/eDespatchIT");
        }

        private void button5_Click_1(object sender, EventArgs e)
        {
            webBrowser1.Navigate("http://webnotes.streamlinefreight.co.nz/");
        }

        private void button7_Click(object sender, EventArgs e)
        {
            Microsoft.Office.Interop.Outlook.Application oApp = new Outlook.Application();
            Outlook._MailItem oMailItem = (Outlook._MailItem)oApp.CreateItem(Outlook.OlItemType.olMailItem);
            oMailItem.To = "26toru@gmail.com";
            // body, bcc etc...
            oMailItem.Display(true);
        }
    }
}
 