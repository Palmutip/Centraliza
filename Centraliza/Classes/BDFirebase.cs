using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using FireSharp.Config;
using FireSharp.Interfaces;
using FireSharp.Response;


namespace Centraliza
{
    public class BDFirebase
    {
        IFirebaseConfig config = new FirebaseConfig
        {
            AuthSecret = "bTq3sR9ZtNDbuqD9K80YZe6Fo9r8B70ZlngYFciY",
            BasePath = "https://mysolar-a8374.firebaseio.com/"
        };
        public IFirebaseClient client;
        public SetResponse setresponse;
        public FirebaseResponse fbresponse;
        public SetResponse setresponse1;
        public FirebaseResponse fbresponse1;
        public SetResponse setresponse2;
        public FirebaseResponse fbresponse2;

        public void FireOpen()
        {
            client = new FireSharp.FirebaseClient(config);
        }
    }
}
