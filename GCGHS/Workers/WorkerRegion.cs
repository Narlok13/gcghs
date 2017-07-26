using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GCGHS.Workers
{
    public class WorkerRegion : Worker
    {
        

        public WorkerRegion(int id, string name, string user, string login, string pass, string otdelName, string otdelNumb, string telOtdel, string ip, string mailWorker, string komment)
            : base (id, name, user, login, pass, otdelName, otdelNumb, telOtdel, ip, mailWorker, komment)
        {
            
        }

        
    }
}
