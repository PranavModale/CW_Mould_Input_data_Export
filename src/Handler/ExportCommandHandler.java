package Handler;

import org.eclipse.core.commands.AbstractHandler;
import org.eclipse.core.commands.ExecutionEvent;

import Service.ExportToExcelService;

public class ExportCommandHandler extends AbstractHandler {

    @Override
    public Object execute(ExecutionEvent event) {
        ExportToExcelService.execute();
        return null;
    }
}
