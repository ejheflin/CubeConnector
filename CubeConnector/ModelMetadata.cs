using System.Collections.Generic;

namespace CubeConnector
{
    public class ModelMetadata
    {
        public List<string> Tables = new List<string>();
        public List<ModelColumn> Columns = new List<ModelColumn>();
        public List<ModelMeasure> Measures = new List<ModelMeasure>();
    }
    public class ModelColumn { public string Table; public string Name; public string DataType; public bool IsHidden;
        public string Qualified => "[" + Table + "].[" + Name + "]"; }
    public class ModelMeasure { public string Table; public string Name; }
}
