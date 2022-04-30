using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Events
{
    public class SQLfunctions
    {
        string adress = @"Data Source=(localdb)\MSSQLLocalDB;Initial Catalog=EventsBD;Integrated Security=True;Connect Timeout=30;Encrypt=False;TrustServerCertificate=False;ApplicationIntent=ReadWrite;MultiSubnetFailover=False";

        public DataTable SelectAllFromEvents()
        {
            DataTable dt = new DataTable();
            SqlConnection con = new SqlConnection(adress);
            SqlCommand cmd = new SqlCommand();
            SqlDataAdapter adapt = new SqlDataAdapter();

            con.Open();
            cmd.CommandText = "Select Events.Id,Events.Name as Название, ConCat(Clients.SurName,' ',Clients.Name,' ', Clients.LastName) as Клиент, ConCat(Responsible.SurName,' ' , Responsible.Name,' ', Responsible.LastName) as Ответственное_лицо, ConCat(TypesOfEvent.Type,' ',Adresses.Name,' ',Adresses.Adress  ) as Помещение, Events.add_inf as ДопИнформация, Events.dateTime From Events Left Join Adresses On Adresses.Id = Events.id_adress inner join TypesOfEvent on Adresses.id_Type = TypesOfEvent.Id Left Join Clients On Clients.Id = Events.id_client Left Join Responsible On Responsible.Id = Events.id_responsible";
            cmd.Connection = con;
            adapt.SelectCommand = cmd;
            adapt.Fill(dt);
            return dt;
        }
        public DataTable SelectInComboboxClient()
        {
            DataTable dt = new DataTable();
            SqlConnection con = new SqlConnection(adress);
            SqlCommand cmd = new SqlCommand();
            SqlDataAdapter adapt = new SqlDataAdapter();
            con.Open();
            cmd.CommandText = "Select Id as Clientid,ConCat(SurName,' ',Name,' ',LastName) as FIO from Clients";
            cmd.Connection = con;
            adapt.SelectCommand = cmd;
            adapt.Fill(dt);
            con.Close();
            return dt;
        }
        public DataTable SelectInComboboxResponsible()
        {
            DataTable dt = new DataTable();
            SqlConnection con = new SqlConnection(adress);
            SqlCommand cmd = new SqlCommand();
            SqlDataAdapter adapt = new SqlDataAdapter();
            con.Open();
            cmd.CommandText = "Select id,ConCat(SurName,' ',Name,' ',LastName) as FIO from Responsible";
            cmd.Connection = con;
            adapt.SelectCommand = cmd;
            adapt.Fill(dt);
            con.Close();
            return dt;
        }
        public DataTable SelectInComboboxAdress(int TypeId)
        {
            DataTable dt = new DataTable();
            SqlConnection con = new SqlConnection(adress);
            SqlCommand cmd = new SqlCommand();
            SqlDataAdapter adapt = new SqlDataAdapter();
            con.Open();
            cmd.CommandText = "Select id,ConCat(Name,' ',Adress) as Adress from Adresses where Adresses.id_Type = " + TypeId;
            cmd.Connection = con;
            adapt.SelectCommand = cmd;
            adapt.Fill(dt);
            con.Close();
            return dt;
        }
        public DataTable SelectInComboboxType()
        {
            DataTable dt = new DataTable();
            SqlConnection con = new SqlConnection(adress);
            SqlCommand cmd = new SqlCommand();
            SqlDataAdapter adapt = new SqlDataAdapter();
            con.Open();
            cmd.CommandText = "Select id,Type from TypesOfEvent";
            cmd.Connection = con;
            adapt.SelectCommand = cmd;
            adapt.Fill(dt);
            con.Close();
            return dt;
        }
        public void AddEvent(string Name, int IdClient, int IdRespons,int IdAdress,string addInf,string dateTime)
        {
            SqlConnection con = new SqlConnection(adress);
            SqlCommand cmd = new SqlCommand();
            cmd.CommandText = "Insert into Events(Name, id_client,id_responsible,id_adress,add_inf,dateTime) Values(N'" + Name + "',N'" + IdClient + "',N'" + IdRespons+ "', N'"+IdAdress+"', N'" + addInf + "',N'" + dateTime + "')";
            cmd.Connection = con;
            con.Open();
            cmd.ExecuteNonQuery();
            con.Close();
        }
    }
}
