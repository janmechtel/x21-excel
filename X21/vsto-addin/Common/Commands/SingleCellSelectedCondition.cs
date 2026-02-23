using X21.Common.Data;
using X21.Common.Model;
using X21.Utils;
using Microsoft.Office.Interop.Excel;

namespace X21.Common.Commands
{
    public class SingleCellSelectedCondition : ModelBase, IExecuteCondition
    {
        public SingleCellSelectedCondition(Container container)
            : base(container)
        {
        }

        public bool CanExecute()
        {
            var result = false;

            result = Execute.Call(() =>
                {
                    var excel = Container.Resolve<Application>();
                    var selection = excel.Selection;

                    if (selection is Range range)
                    {
                        return range.Cells.Count == 1;
                    }

                    return false;
                },
                Execute.CatchMode.DontLog
            );

            return result;
        }
    }
}
