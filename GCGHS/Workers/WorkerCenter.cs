using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GCGHS.Workers
{
    public class WorkerCenter : Worker
    {
        string telVnutr;
        string roomNumber;

        public WorkerCenter(int id, string name, string user, string login, string pass, string otdelName, string otdelNumb, string telOtdel, string telVnutr, string ip, string mailWorker, string komment, string roomNumber)
            : base(id, name, user, login, pass, otdelName, otdelNumb, telOtdel, ip, mailWorker, komment)
        {
            
            TelVnutr = telVnutr;
            RoomNumber = roomNumber;
        }

        public string TelVnutr { get => telVnutr; set => telVnutr = value; }
        public string RoomNumber { get => roomNumber; set => roomNumber = value; }
    }
}
