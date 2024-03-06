from openpyxl import Workbook
from openpyxl_toolbox.sensitivity_manager import (
    create_sensitivity_label_definition,
    get_label_from_file,
    set_label_to_workbook,
    MSIP_Configuration,
)

if __name__ == "__main__":
    # use this line to create a configuration json file from a series of models
    """
    default value
    DEFAULT_SENSITIVITY_LABELS_DEFINITION = "sensitivity_model/sensitivity_labels_definition_with_openpyxl.json"
    DEFAULT_SENSITIVITY_TEMPLATES = "sensitivity_model"
    """
    msip_configuration = create_sensitivity_label_definition()

    # use this to load the json configuration
    msip_configuration = MSIP_Configuration().load()

    for label_name in msip_configuration.labels():
        test_filename = f"output/dummy_openpyxl_{label_name}.xlsx"
        wb = Workbook()
        label = msip_configuration.get_sensitivity_label(label_name)
        set_label_to_workbook(wb, label)
        wb.save(test_filename)
        print(
            f"""
{test_filename} is created with a Sensitivity Label ({label_name})"""
        )

        check_label = get_label_from_file(test_filename)
        assert check_label is not None
        assert label.LabelId == check_label.LabelId
        assert label.LabelName == check_label.LabelName
        print(
            f"{test_filename} is correctly assigned a Sensitivity Label ({label_name})"
        )
