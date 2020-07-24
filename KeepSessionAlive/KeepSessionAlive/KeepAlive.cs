using RdcMan;
using System.Windows.Forms;
using System.Xml;
using System.ComponentModel.Composition;

namespace RdcManPlugin
{
    [Export(typeof(IPlugin))]
    public class KeepSessionAlive : IPlugin
    {
        public void OnContextMenu(System.Windows.Forms.ContextMenuStrip contextMenuStrip, RdcTreeNode node)
        {
            throw new System.NotImplementedException();
        }

        public void OnDockServer(ServerBase server)
        {
            throw new System.NotImplementedException();
        }

        public void OnUndockServer(IUndockedServerForm form)
        {
            throw new System.NotImplementedException();
        }

        public void PostLoad(IPluginContext context)
        {
            throw new System.NotImplementedException();
        }

        public void PreLoad(IPluginContext context, XmlNode xmlNode)
        {
            throw new System.NotImplementedException();
        }

        public XmlNode SaveSettings()
        {
            throw new System.NotImplementedException();
        }

        public void Shutdown()
        {
            throw new System.NotImplementedException();
        }
    }
}
