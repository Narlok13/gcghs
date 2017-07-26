using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GCGHS.Workers
{
    abstract public class Worker
    {
        string name;
        string user;
        string login;
        string pass;
        string otdelName;
        string telOtdel;
        string otdelNumb;
        string ip;
        int id;
        string mailWorker;
        string komment;

        public Worker(int id, string name, string user, string login, string pass, string otdelName, string otdelNumb, string telOtdel, string ip, string mailWorker, string komment)
        {
            Name = name;
            User = user;
            Login = login;
            Pass = pass;
            OtdelName = otdelName;
            TelOtdel = telOtdel;
            OtdelNumb = otdelNumb;
            Ip = ip;
            Id = id;
            MailWorker = mailWorker;
            Komment = komment;
        }

        public string Name { get => name; set => name = value; }
        public string User { get => user; set => user = value; }
        public string Login { get => login; set => login = value; }
        public string Pass { get => pass; set => pass = value; }
        public string OtdelName { get => otdelName; set => otdelName = value; }
        public string TelOtdel { get => telOtdel; set => telOtdel = value; }
        public string OtdelNumb { get => otdelNumb; set => otdelNumb = value; }
        public string Ip { get => ip; set => ip = value; }
        public int Id { get => id; set => id = value; }
        public string MailWorker { get => mailWorker; set => mailWorker = value; }
        public string Komment { get => komment; set => komment = value; }

        public void SendTask()
        {

        }
    }
}
