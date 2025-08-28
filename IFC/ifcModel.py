import ifcopenshell
import pandas as pd
import json
import ifcopenshell.util
import ifcopenshell.util.element

# Load IFC model
#model = ifcopenshell.open("model.ifc")
model = ifcopenshell.open("Infra-Bridge.ifc")

# =============================================================================
# 
# all_classes= model.types()
# 
# #print(len(all_classes))
# 
# schema= model.schema
# 
# print(f"schema={schema}")
# 
# # Get the project
# project = model.by_type("IfcProject")[0]
# 
# # Get owner history
# owner_history = project.OwnerHistory
# 
# # Extract author name
# author_person = owner_history.OwningUser.ThePerson
# author_name = f"{author_person.GivenName} {author_person.FamilyName}"
# 
# print("Author:", author_name)
# 
# beams= model.by_type("IfcBeam")
# beam= beams[0]
# 
# print(beam.is_a())# print the class name
# 
# print(beam.Name) # name of atribuite
# 
# #print(beam.__class__)
# 
# print(f" all related properties: {ifcopenshell.util.element.get_psets(beam)}_________________")
# 
# print(model.get_inverse(beam))
# 
# print("_____________________")
# 
# print(model.traverse(beam, max_levels=1))
# 
# 
# print(beams[-1])
# 
# for b in beams:
# 	print(f"{b.GlobalId}: {b.Name}______________")
# 	
# 
# #beam0=beams[0]
# 
# #beam0.Name= "{b.Name} +  New "
# 
# for b in beams:
# 	print(f"{b.GlobalId}: {b.Name}")
# 	b.Name= f"{b.Name} "
# 	print(f"{b.GlobalId}: {b.Name}")
# 	
# print("----------------------------------------")
# 
# for b in beams:
# 	print(f"{b.GlobalId}: {b.Name}")
# 
# 
# model.create_entity('IfcDoor',ifcopenshell.guid.new(), Name= "Door for Bridge")
# 
# #print(model.by_type("IfcDoor"))
# 
# all_cls= model.types()
# print(all_cls)
# 
# =============================================================================

# Dictionary to group objects by Name
objects_by_name = {}

for element in model.by_type("IfcObject"):
	element_data = {}

	# Basic info
	info = element.get_info()
	for key, value in info.items():
		if key not in ["id", "OwnerHistory", "Tag", "PredefinedType"]:
			element_data[key] = value

	# Ensure Name exists (fallback to GlobalId if missing)
	object_name = info.get("Name") or element.GlobalId
	element_data["Name"] = object_name

	 # Materials
	materials = []
	if hasattr(element, "HasAssociations"):
		for rel in element.HasAssociations:
			if rel.is_a("IfcRelAssociatesMaterial"):
				mat = rel.RelatingMaterial
				if mat:
					# Handle composite materials
					if hasattr(mat, "Materials"):
						materials.extend([m.Name for m in mat.Materials])
					else:
						materials.append(mat.Name)
	element_data["Materials"] = ", ".join(materials) if materials else None

	# Quantities
	if hasattr(element, "IsDefinedBy"):
		for rel in element.IsDefinedBy:
			pset = rel.RelatingPropertyDefinition
			if pset.is_a("IfcElementQuantity"):
				for quantity in pset.Quantities:
					qname = f"Quantity_{quantity.Name}"
					qvalue = None
					if quantity.is_a("IfcQuantityLength"):
						qvalue = quantity.LengthValue
					elif quantity.is_a("IfcQuantityArea"):
						qvalue = quantity.AreaValue
					elif quantity.is_a("IfcQuantityVolume"):
						qvalue = quantity.VolumeValue
					elif hasattr(quantity, "NominalValue"):
						qvalue = getattr(quantity.NominalValue, "wrappedValue", None)
					element_data[qname] = qvalue

	# Property sets
	if hasattr(element, "IsDefinedBy"):
		for rel in element.IsDefinedBy:
			pset = rel.RelatingPropertyDefinition
			if pset.is_a("IfcPropertySet"):
				for prop in pset.HasProperties:
					prop_name = f"{pset.Name}_{prop.Name}"
					value = None
					if prop.is_a("IfcPropertySingleValue") and prop.NominalValue:
						value = prop.NominalValue.wrappedValue
					elif prop.is_a("IfcPropertyEnumeratedValue"):
						value = [v.wrappedValue for v in getattr(prop, "EnumerationValues", []) if v]
					elif prop.is_a("IfcPropertyBoundedValue") and prop.UpperBoundValue:
						value = prop.UpperBoundValue.wrappedValue
					elif prop.is_a("IfcPropertyListValue"):
						value = [v.wrappedValue for v in getattr(prop, "ListValues", []) if v]
					element_data[prop_name] = value

	element_data.pop("TypeDescription", None)

	# Group by Name
	if object_name not in objects_by_name:
		objects_by_name[object_name] = []
	objects_by_name[object_name].append(element_data)
	# sort object_name alphabetically
	objects_by_name = {k: objects_by_name[k] for k in sorted(objects_by_name)}
	

# --------------------------
# Export each group to its own sheet
# --------------------------
output_file = "Infra-Bridge_by_grouped_name.xlsx"
#output_file = "Building-Architecture.xlsx"
with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
	for name, elements in objects_by_name.items():
		safe_sheet_name = str(name)[:31]  # Excel sheet name limit
		df = pd.DataFrame(elements)  # multiple rows for same Name
		df.to_excel(writer, sheet_name=safe_sheet_name, index=False)

print(f"✅ Exported {len(objects_by_name)} grouped sheets to {output_file}")


#%% Group by class name

# =============================================================================
# 
# import ifcopenshell
# import pandas as pd
# from collections import Counter
# import ifcopenshell.util.selector  # optional
# 
# # --------------------------
# # Part 1: Load model & collect data
# # --------------------------
# 
# # Load IFC model
# model = ifcopenshell.open("Infra-Bridge.ifc")
# 
# # Get all class names
# class_names = sorted(model.types())
# print("\nTotal classes found:", len(class_names))
# 
# # Print class names
# for i, c in enumerate(class_names):
#  print(f"{i}: {c}")
# 
# # Collect element info by class
# all_elements = {}  # dictionary: class_name -> list of elements info
# 
# for cls in class_names:
#  elements = model.by_type(cls)
#  print(f"{cls}: {len(elements)} elements")
#  if not elements:
# 	 continue
# 
#  rows = []
#  for element in elements:
# 	 element_data = {}
# 
# 	 # Basic info
# 	 info = element.get_info()
# 	 for key, value in info.items():
# 		 if key not in ["id", "OwnerHistory"]:
# 			 element_data[key] = value
# 
# 	 # If element has a type
# 	 type_obj = None
# 	 if hasattr(element, "IsTypedBy") and element.IsTypedBy:
# 		 type_obj = element.IsTypedBy[0].RelatingType
# 
# 	 #element_data["TypeName"] = type_obj.Name if type_obj else None
# 	 element_data["TypeDescription"] = type_obj.Description if type_obj else None
# 
# 	 # Materials
# 	 materials = []
# 	 if hasattr(element, "HasAssociations"):
# 		 for rel in element.HasAssociations:
# 			 if rel.is_a("IfcRelAssociatesMaterial"):
# 				 mat = rel.RelatingMaterial
# 				 if mat:
# 					 # Handle composite materials
# 					 if hasattr(mat, "Materials"):
# 						 materials.extend([m.Name for m in mat.Materials])
# 					 else:
# 						 materials.append(mat.Name)
# 	 element_data["Materials"] = ", ".join(materials) if materials else None
# 
# 	 # Quantities
# 	 if hasattr(element, "IsDefinedBy"):
# 		 for rel in element.IsDefinedBy:
# 			 pset = rel.RelatingPropertyDefinition
# 			 if pset.is_a("IfcElementQuantity"):
# 				 for quantity in pset.Quantities:
# 					 qname = f"Quantity_{quantity.Name}"
# 					 qvalue = quantity.get_info().get("NominalValue")
# 					 element_data[qname] = qvalue
# 
# 	 # Property sets
# 	 if hasattr(element, "IsDefinedBy"):
# 		 for rel in element.IsDefinedBy:
# 			 pset = rel.RelatingPropertyDefinition
# 			 if pset.is_a("IfcPropertySet"):
# 				 for prop in pset.HasProperties:
# 					 prop_name = f"{pset.Name}_{prop.Name}"
# 					 value = None
# 
# 					 # Handle different property types
# 					 if prop.is_a("IfcPropertySingleValue"):
# 						 value = getattr(prop.NominalValue, "wrappedValue", None)
# 					 elif prop.is_a("IfcPropertyEnumeratedValue"):
# 						 value = [v.wrappedValue for v in getattr(prop, "EnumerationValues", [])]
# 					 elif prop.is_a("IfcPropertyBoundedValue"):
# 						 value = getattr(prop, "UpperBoundValue", None)
# 					 elif prop.is_a("IfcPropertyListValue"):
# 						 value = [v.wrappedValue for v in getattr(prop, "ListValues", [])]
# 
# 					 element_data[prop_name] = value
# 
# 	 rows.append(element_data)
# 
#  all_elements[cls] = rows
# 
# # --------------------------
# # Part 2: Export to Excel with ClassTypes sheet
# # --------------------------
# 
# with pd.ExcelWriter("Infra-Bridge_full_info_with_links.xlsx", engine="xlsxwriter") as writer:
#  workbook = writer.book
# 
#  # 1️⃣ Create main sheet "ClassTypes"
#  class_types_list = []
#  for cls_name in sorted(all_elements.keys()):
# 	 sheet_name = cls_name if len(cls_name) <= 31 else cls_name[:28] + "..."
# 	 class_types_list.append({"ClassName": cls_name, "SheetName": sheet_name})
# 
#  df_class_types = pd.DataFrame(class_types_list)
#  df_class_types.to_excel(writer, sheet_name="ClassTypes", index=False)
#  worksheet_main = writer.sheets["ClassTypes"]
# 
#  # Add hyperlinks
#  for row_num, cls_info in enumerate(class_types_list, start=1):  # start=1 for header
# 	 worksheet_main.write_url(
# 		 f"A{row_num + 1}",
# 		 f"internal:'{cls_info['SheetName']}'!A1",
# 		 string=cls_info['ClassName']
# 	 )
# 
#  # 2️⃣ Create individual sheets per class
#  for cls_name, rows in all_elements.items():
# 	 if not rows:
# 		 continue
# 	 sheet_name = cls_name if len(cls_name) <= 31 else cls_name[:28] + "..."
# 	 df = pd.DataFrame(rows)
# 	 df.to_excel(writer, sheet_name=sheet_name, index=False)
# 
# print("✅ Exported all classes with hyperlinks to Infra-Bridge_full_info_with_links.xlsx")
# 
# =============================================================================


