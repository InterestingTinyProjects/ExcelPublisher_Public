using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WebSocket.Client.Netty
{
    class WebSocketClient
    {

        public WebSocketClient()
        {
            var builder = new UriBuilder
            {
                Scheme = "wss",
                Host = "gateway.saxobank.com",
                Port = 80,
                Path = "/sim/openapi/streamingws/connect"
            };
            var uri = builder.Uri;


        }

    }
}
