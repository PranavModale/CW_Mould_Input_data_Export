package Service;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.nio.file.Files;
import java.nio.file.StandardCopyOption;
import java.util.HashMap;
import java.util.Map;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.eclipse.jface.dialogs.MessageDialog;
import org.eclipse.swt.widgets.Display;
import org.eclipse.swt.widgets.Shell;

import com.teamcenter.rac.aif.kernel.InterfaceAIFComponent;
import com.teamcenter.rac.aifrcp.AIFUtility;
import com.teamcenter.rac.kernel.TCComponentDataset;
import com.teamcenter.rac.kernel.TCComponentItemRevision;

import Util.DatasetUtil;
import Util.ExcelUtil;

public class ExportToExcelService {

    private static final String TEMPLATE_DATASET_NAME = "Input Data Sheet for Mould Design";
    private static final String OUTPUT_DATASET_NAME = "Input Data Sheet for Mould Design.xlsx";
    private static final String ALLOWED_OBJECT_TYPE = "Mould Assembly Revision";

    public static void execute() {

        FileInputStream fis = null;
        FileOutputStream fos = null;
        XSSFWorkbook workbook = null;

        try {
            InterfaceAIFComponent selected =
                    AIFUtility.getCurrentApplication().getTargetComponent();

            if (!(selected instanceof TCComponentItemRevision)) {
                showError(
                        "Export Mould Data to Excel",
                        "Please select an Item Revision that contains the dataset \""
                                + TEMPLATE_DATASET_NAME + "\".");
                return;
            }

            TCComponentItemRevision itemRev = (TCComponentItemRevision) selected;
            
            String objectType = itemRev.getProperty("object_type");
            if (objectType == null || !ALLOWED_OBJECT_TYPE.equalsIgnoreCase(objectType.trim())) {
                showError(
                        "Export Mould Data to Excel",
                        "Please select object type \"" + ALLOWED_OBJECT_TYPE + "\".");
                return;
            }

            TCComponentDataset dataset = DatasetUtil.findDataset(itemRev, TEMPLATE_DATASET_NAME);
            if (dataset == null) {
                showError(
                        "Export Mould Data to Excel",
                        "Selected object does not contain the dataset \""
                                + TEMPLATE_DATASET_NAME
                                + "\".\n\nPlease select the correct revision.");
                return;
            }

            File templateFile = DatasetUtil.downloadDataset(dataset);

            fis = new FileInputStream(templateFile);
            workbook = new XSSFWorkbook(fis);
            fis.close();
            fis = null;

            Sheet sheet = workbook.getSheetAt(0);

            itemRev.refresh();

            ExcelUtil.processSheet(sheet, itemRev, getMapping());
            ExcelUtil.processTableBlocks(sheet, itemRev, getTableMapping());

            File outDir = new File("C:\\Temp");
            if (!outDir.exists()) {
                outDir.mkdirs();
            }

            File tempOutput = new File(
                    outDir,
                    "MouldData_temp_" + System.currentTimeMillis() + ".xlsx");
            File finalOutput = new File(outDir, OUTPUT_DATASET_NAME);

            fos = new FileOutputStream(tempOutput);
            workbook.write(fos);
            fos.close();
            fos = null;

            workbook.close();
            workbook = null;

            Files.move(
                    tempOutput.toPath(),
                    finalOutput.toPath(),
                    StandardCopyOption.REPLACE_EXISTING);

            if (!templateFile.getAbsolutePath().equalsIgnoreCase(finalOutput.getAbsolutePath())) {
                try {
                    Files.deleteIfExists(templateFile.toPath());
                } catch (Exception ex) {
                    System.out.println(
                            "Warning: could not delete old downloaded template: "
                                    + templateFile.getAbsolutePath());
                }
            }

            DatasetUtil.uploadExcelToInformation(itemRev, finalOutput, OUTPUT_DATASET_NAME);

            try {
                Files.deleteIfExists(finalOutput.toPath());
            } catch (Exception ex) {
                System.out.println(
                        "Warning: uploaded successfully, but could not delete local file: "
                                + finalOutput.getAbsolutePath());
            }

            MessageDialog.openInformation(
                    getShell(),
                    "Export Mould Data to Excel",
                    "Excel exported and uploaded successfully.");

        } catch (Exception e) {
            e.printStackTrace();
            showError(
                    "Export Mould Data to Excel",
                    "Export failed.\n\n" + e.getMessage());

        } finally {
            try {
                if (fis != null) {
                    fis.close();
                }
            } catch (Exception ignored) {
            }

            try {
                if (fos != null) {
                    fos.close();
                }
            } catch (Exception ignored) {
            }

            try {
                if (workbook != null) {
                    workbook.close();
                }
            } catch (Exception ignored) {
            }
        }
    }

    private static void showError(String title, String message) {
        MessageDialog.openError(getShell(), title, message);
    }

    private static Shell getShell() {
        if (Display.getDefault() == null) {
            return null;
        }
        return Display.getDefault().getActiveShell();
    }

    public static Map<String, String> getTableMapping() {

        Map<String, String> map = new HashMap<String, String>();

        // Mould Base
        map.put("Mould Base|Insulation Sheet|Material", "a2InsulationSheetMaterial");
        map.put("Mould Base|Insulation Sheet|Specific Vendor", "a2InsulationSheetSpecVendor");
        map.put("Mould Base|Insulation Sheet|Specific Heat Treatment process", "a2InsulationSheetSHTProcess");
        map.put("Mould Base|Insulation Sheet|Hardness/Remark", "a2InsulationSheetHRemark");

        map.put("Mould Base|Top Plate|Material", "a2TopPlateMaterial");
        map.put("Mould Base|Top Plate|Specific Vendor", "a2TopPlateSpecificVendor");
        map.put("Mould Base|Top Plate|Specific Heat Treatment process", "a2TopPlateSpecHeatProcess");
        map.put("Mould Base|Top Plate|Hardness/Remark", "a2TopPlateHardnessRemark");

        map.put("Mould Base|Manifold Plate|Material", "a2ManifoldPlateMaterial_New");
        map.put("Mould Base|Manifold Plate|Specific Vendor", "a2ManifoldPlateSpecVendor");
        map.put("Mould Base|Manifold Plate|Specific Heat Treatment process", "a2ManifoldPlateSHTProcess");
        map.put("Mould Base|Manifold Plate|Hardness/Remark", "a2ManifoldPlateHRemark");

        map.put("Mould Base|Cavity Plate|Material", "a2CavityPlateMaterial");
        map.put("Mould Base|Cavity Plate|Specific Vendor", "a2CavityPlateSpecificVendor");
        map.put("Mould Base|Cavity Plate|Specific Heat Treatment process", "a2CavityPlateSHTProcess");
        map.put("Mould Base|Cavity Plate|Hardness/Remark", "a2CavityPlateHardnessRemark");

        map.put("Mould Base|Core Plate|Material", "a2CorePlateMaterial");
        map.put("Mould Base|Core Plate|Specific Vendor", "a2CorePlateSpecificVendor");
        map.put("Mould Base|Core Plate|Specific Heat Treatment process", "a2CorePlateSHTProcess");
        map.put("Mould Base|Core Plate|Hardness/Remark", "a2CorePlateHardnessRemark");

        map.put("Mould Base|Stripper Plate|Material", "a2StripperPlateMaterial");
        map.put("Mould Base|Stripper Plate|Specific Vendor", "a2StripperPlateSpecVendor");
        map.put("Mould Base|Stripper Plate|Specific Heat Treatment process", "a2StripperPlateSHTProcess");
        map.put("Mould Base|Stripper Plate|Hardness/Remark", "a2StripperPlateHRemark");

        map.put("Mould Base|Back Plates(if required)|Material", "a2BackPlatesMaterial");
        map.put("Mould Base|Back Plates(if required)|Specific Vendor", "a2BackPlatesSpecificVendor");
        map.put("Mould Base|Back Plates(if required)|Specific Heat Treatment process", "a2BackPlatesSHTProcess");
        map.put("Mould Base|Back Plates(if required)|Hardness/Remark", "a2BackPlatesHardnessRemark");

        map.put("Mould Base|Ejector Plate|Material", "a2EjectorPlateMaterial");
        map.put("Mould Base|Ejector Plate|Specific Vendor", "a2EjectorPlateSpecVendor");
        map.put("Mould Base|Ejector Plate|Specific Heat Treatment process", "a2EjectorPlateSHTProcess");
        map.put("Mould Base|Ejector Plate|Hardness/Remark", "a2EjectorPlateHRemark");

        map.put("Mould Base|Ejector Back Plate|Material", "a2EjectorBackPlateMaterial");
        map.put("Mould Base|Ejector Back Plate|Specific Vendor", "a2EjectorBackPlateSpeVendor");
        map.put("Mould Base|Ejector Back Plate|Specific Heat Treatment process", "a2EjectorBackPlateSHTP");
        map.put("Mould Base|Ejector Back Plate|Hardness/Remark", "a2EjectorBackPlateHRemark");

        map.put("Mould Base|Base Plate|Material", "a2BasePlateMaterial");
        map.put("Mould Base|Base Plate|Specific Vendor", "a2BasePlateSpecificVendor");
        map.put("Mould Base|Base Plate|Specific Heat Treatment process", "a2BasePlateSHTProcess");
        map.put("Mould Base|Base Plate|Hardness/Remark", "a2BasePlateHardnessRemark");

        map.put("Mould Base|Parallel Block|Material", "a2ParallelBlockMaterial");
        map.put("Mould Base|Parallel Block|Specific Vendor", "a2ParallelBlockSpecVendor");
        map.put("Mould Base|Parallel Block|Specific Heat Treatment process", "a2ParallelBlockSHTProcess");
        map.put("Mould Base|Parallel Block|Hardness/Remark", "a2ParallelBlockHRemark");

        // Insert Assembly
        map.put("Insert Assembly|Core Insert|Material", "a2CoreInsertMaterial");
        map.put("Insert Assembly|Core Insert|Specific Vendor", "a2CoreInsertSpecificVendor");
        map.put("Insert Assembly|Core Insert|Specific Heat Treatment Stroke Processes", "a2CoreInsertSHTProcess");
        map.put("Insert Assembly|Core Insert|Hardness/Remark", "a2CoreInsertHardnessRemark");

        map.put("Insert Assembly|Cavity Insert|Material", "a2CavityInsertMaterial");
        map.put("Insert Assembly|Cavity Insert|Specific Vendor", "a2CavityInsertSpecVendor");
        map.put("Insert Assembly|Cavity Insert|Specific Heat Treatment Stroke Processes", "a2CavityInsertSHTProcess");
        map.put("Insert Assembly|Cavity Insert|Hardness/Remark", "a2CavityInsertHRemark");

        map.put("Insert Assembly|Moving Core|Material", "a2MovingCoreMaterial");
        map.put("Insert Assembly|Moving Core|Specific Vendor", "a2MovingCoreSpecificVendor");
        map.put("Insert Assembly|Moving Core|Specific Heat Treatment Stroke Processes", "a2MovingCoreSHTProcess");
        map.put("Insert Assembly|Moving Core|Hardness/Remark", "a2MovingCoreHardnessRemark");

        map.put("Insert Assembly|Sub-inserts|Material", "a2SubInsertsMaterial");
        map.put("Insert Assembly|Sub-inserts|Specific Vendor", "a2SubInsertsSpecificVendor");
        map.put("Insert Assembly|Sub-inserts|Specific Heat Treatment Stroke Processes", "a2SubInsertsSHTProcess");
        map.put("Insert Assembly|Sub-inserts|Hardness/Remark", "a2SubInsertsHardnessRemark");

        map.put("Insert Assembly|Engraving Inserts|Material", "a2EngravingInsertsMaterial");
        map.put("Insert Assembly|Engraving Inserts|Specific Vendor", "a2EngravingInsertsSpcVendor");
        map.put("Insert Assembly|Engraving Inserts|Specific Heat Treatment Stroke Processes", "a2EngravingISHTSProcess");
        map.put("Insert Assembly|Engraving Inserts|Hardness/Remark", "a2EngravingInsertsHRemark");

        map.put("Insert Assembly|Cut-Bearing Inserts|Material", "a2CutBearingInsertsMaterial");
        map.put("Insert Assembly|Cut-Bearing Inserts|Specific Vendor", "a2CutBearingInsertSpcVendor");
        map.put("Insert Assembly|Cut-Bearing Inserts|Specific Heat Treatment Stroke Processes", "a2CutBearingISHTSProcesses");
        map.put("Insert Assembly|Cut-Bearing Inserts|Hardness/Remark", "a2CutBearingInsertsHRemark");

        return map;
    }

    public static Map<String, String> getMapping() {

        Map<String, String> map = new HashMap<String, String>();

        // Input Data Sheet For Mould Design
        map.put("Project No", "a2ProjectNo");
        map.put("Project Name", "a2CatalogueNoComp_P");
        map.put("Sub- Project No.", "a2SubProjectNo");
        map.put("Sub-Project Name", "a2SubProjectName");
        map.put("Manufacturer", "a2Manufacturer");
        map.put("Project Type", "a2ProjectType");
        map.put("Date", "a2Date");
        map.put("Prepared By", "a2PreparedBy");

        // PART
        map.put("Part Name", "a2PartNameDPP");
        map.put("PLM ID/Drawing No. & Revision", "MULTI_PLM_DRAWING_REV");
        map.put("Part Weight", "a2PartWeightDPP");
        map.put("Raw Material", "a2RawMaterialDPP");
        map.put("Raw Material Grade", "a2RawMaterialGradeDPP");
        map.put("Shrinkage", "a2RawMaterialShrinkage");
        map.put("Colour", "a2ColorDPP");
        map.put("Surface Finish", "a2SurfaceFinish");
        map.put("Version Change-Over(If any)", "a2VersionChangeOverNotePart");
        map.put("Additional Note (If any)", "a2AdditionalNotePart");

        // MOULD
        map.put("Mould Type", "a2MouldType");
        map.put("No. of Cavity", "a2NoOfCavity");
        map.put("Expected tool Life", "a2ExpectedToolLife_New");
        map.put("Version Change Over Note", "a2VersionChngeOverNoteMould");
        map.put("Additional note (If any)", "a2AdditionalNoteMould");
        map.put("Sprue Type", "a2SprueType");
        map.put("Runner Type", "a2TypeofRunner");
        map.put("Injection Flow Path", "a2InjectionFlowPath");
        map.put("Gate", "a2GatetypeNew");
        map.put("HRS OEM", "a2HRS_OEM");
        map.put("No. of drops for Semi / Hot Runner System", "a2NoOfDrops_SemiHRS_OR_HRS");
        map.put("Moulding Machine", "a2TypeofmoldingmachineNew");

        // Mould Base
        map.put("Std. Boughtout items for Mould Base( eg guiding elements , interlock , Return Pins etc.)", "a2StdBoughtoutItemMouldBase");

        // Insert Assembly
        map.put("Ejector Pin", "a2EjectorPin");
        map.put("Ejector Pin Type", "a2EjectorPinType");
        map.put("Ejector Pin make", "a2EjectorPinMake");
        map.put("Numbering pin/Coring pin", "a2NumberingOrCoringPin");
        map.put("Centering Sleeve", "a2CenteringSleeve");
        map.put("CW Standard Moving Core Units", "a2CWStandardMovingCoreUnits");

        // Cooling Accessories
        map.put("Cooling In/Out Position", "a2CoolingInOutPosition");

        // Air vent
        map.put("Air vent need to provide as per the Moldex Analysis", "a2AirVentMoldexAnalysis");

        // Design (CAD Format)
        map.put("For In-house", "a2CADDesignFormatForInhouse");
        map.put("For Outsource", "a2CADDesignFormatOutsource");

        return map;
    }
}