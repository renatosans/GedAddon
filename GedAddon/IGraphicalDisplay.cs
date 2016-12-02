using System;
using System.Collections.Generic;


namespace SystemIntegration
{
    public interface IGraphicalDisplay
    {
        void NotifyEvent(Object eventObject);
        Object ShowInfo(String information);
        void CloseInfo(Object infoDialog);
        Object PrepareGrid(String gridOwner, String gridName, List<Object> columns, int rowCount);
        void AddGridRow(Object grid, int rowIndex, Object[] values);
    }

}
