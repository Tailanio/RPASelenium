using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System;
using System.Collections.ObjectModel;
using System.IO;
using System.Data;
using ExcelDataReader;
using System.Text;

namespace ConsoleSelenium
{
    struct Infos {
        public string firstName;
        public string lastName;
        public string companyName;
        public string role;
        public string address;
        public string email;
        public double phone;

        public string toString()
        {
            return firstName + " " + lastName + "" + companyName + " " + role + " " + address + " " + email + " " + phone;
        }
    }


    class Program
    {
        static void Main(string[] args)
        { 
            //carrega os dados da tabela do excel
            Collection<Infos> lista = carregarData();

            //conecta o driver do chrome com selenium 
            IWebDriver driver = new ChromeDriver(Directory.GetCurrentDirectory());
            driver.Url = "https://www.rpachallenge.com/";

            //identifica o botao start do challenge
            IWebElement btnStart = driver.FindElement(By.TagName("button"));
            btnStart.Click();

            //identifica o botao de submit 
            IWebElement btnSubmit = driver.FindElement(By.XPath("/html/body/app-root/div[2]/app-rpa1/div/div[2]/form/input"));

            //percorre a lista de entradas e preenche com os dados
            foreach (Infos i in lista) {
                ReadOnlyCollection<IWebElement> inputs = driver.FindElements(By.TagName("input"));
                foreach (WebElement e in inputs)
                {
                    switch (e.GetAttribute("ng-reflect-name"))
                    {
                        case "labelPhone":
                            e.SendKeys(i.phone.ToString());
                            break;

                        case "labelAddress":
                            e.SendKeys(i.address);
                            break;

                        case "labelCompanyName":
                            e.SendKeys(i.companyName);
                            break;

                        case "labelEmail":
                            e.SendKeys(i.email);
                            break;

                        case "labelRole":
                            e.SendKeys(i.role);
                            break;

                        case "labelFirstName":
                            e.SendKeys(i.firstName);
                            break;

                        case "labelLastName":
                            e.SendKeys(i.lastName);
                            break;

                        default:
                            break;
                    }
                }
                if (btnSubmit != null) {
                    btnSubmit.Click();
                }
            }
        }

        private static Collection<Infos> carregarData()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            FileStream fStream = File.Open(Directory.GetCurrentDirectory() + @"\challenge.xlsx", FileMode.Open, FileAccess.Read);
            var excelDataReader = ExcelReaderFactory.CreateReader(fStream);

            var resultDataSet = excelDataReader.AsDataSet();

            var rows = resultDataSet.Tables[0].Rows;

            Collection<Infos> lista = new Collection<Infos>();

            foreach (DataRow r in rows)
            {
                if (r == rows[0])
                    continue;
                Infos info = new Infos();
                info.firstName = r.Field<string>(0);
                info.lastName = r.Field<string>(1);
                info.companyName = r.Field<string>(2);
                info.role = r.Field<string>(3);
                info.address = r.Field<string>(4);
                info.email = r.Field<string>(5);
                info.phone = r.Field<double>(6);

                Console.WriteLine(info.toString() + '\n');
                lista.Add(info);
            }

            excelDataReader.Close();
            return lista;
        }
    }
}
