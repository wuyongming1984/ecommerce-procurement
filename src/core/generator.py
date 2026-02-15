import os
from pathlib import Path
from .template_mgr import TemplateManager
import jinja2

class DocumentGenerator:
    def __init__(self, template_dir):
        self.template_mgr = TemplateManager(template_dir)

    def generate_batch(self, data_list, common_context, templates, output_dir, name_field="Name"):
        """
        Generates documents for each item in data_list.
        
        Args:
            data_list (list): List of dictionaries containing row data.
            common_context (dict): Dictionary of common fields (Public Info).
            templates (list): List of template filenames to render for each row.
            output_dir (str): Directory to save generated files.
            name_field (str): Field in data_list to use for naming output files.
        
        Returns:
            list: List of result dictionaries.
        """
        results = []
        output_path_root = Path(output_dir)
        output_path_root.mkdir(parents=True, exist_ok=True)

        for i, item_data in enumerate(data_list):
            # Merge contexts: Item data overrides common data if conflict (or vice versa? usually Item specific > Common)
            # Actually, let's merge: Common base, update with Item.
            context = common_context.copy()
            context.update(item_data)
            
            # Determine a base name for files
            base_name = str(item_data.get(name_field, f"Item_{i+1}"))
            # Clean base_name of invalid characters if necessary
            base_name = "".join([c for c in base_name if c.isalpha() or c.isdigit() or c in (' ', '-', '_')]).strip()

            item_results = {"item": base_name, "files": [], "errors": []}

            for template_name_raw in templates:
                try:
                    # Render template filename (e.g. "Draft_{{Type}}.docx" -> "Draft_A.docx")
                    template_name = jinja2.Template(template_name_raw).render(context)
                    print(f"DEBUG: Template resolve: '{template_name_raw}' -> '{template_name}'")
                    
                    # Construct output filename: {BaseName}_{TemplateName}
                    # Remove .docx/.doc from template name for the suffix
                    template_path_obj = Path(template_name)
                    template_stem = template_path_obj.stem
                    output_filename = f"{base_name}_{template_stem}.docx"
                    output_full_path = output_path_root / output_filename
                    
                    self.template_mgr.render_and_save(
                        template_name,
                        context,
                        output_full_path
                    )
                    item_results["files"].append(str(output_full_path))
                except Exception as e:
                    item_results["errors"].append(f"Template {template_name}: {str(e)}")
            
            results.append(item_results)
            
        return results
