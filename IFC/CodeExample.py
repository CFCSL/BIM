import ifcopenshell

# load model
model = ifcopenshell.open('Infra-Bridge.ifc')

#%%
# 1. Model Inspection & Metadata Extraction

# Schema
print(model.schema)

# Project info
project = model.by_type("IfcProject")[0]
print("Project Name:", project.Name)

# Author info
author_name = f"{project.OwnerHistory.OwningUser.ThePerson.GivenName} {project.OwnerHistory.OwningUser.ThePerson.FamilyName}"
print("Author:", author_name)

# Get all entites (STEP ID items)
print(f"There are {len(model.types())} STEP entities, which are: ") 
for inst in model.types():
	# It prints out entire entity as it appears in IFC file, e.g
	print( inst)
	
print ("_________________________")

#%%
#2. Geometry & Spatial Analysis
import ifcopenshell.util.placement
import ifcopenshell.geom
import ifcopenshell.util.shape

#Create geometric shape for an IfcBeam:
beam = model.by_type("IfcBeam")[0]
settings = ifcopenshell.geom.settings()
shape = ifcopenshell.geom.create_shape(settings, beam)

# The GUID of the element we processed
print(shape.guid)

# The ID of the element we processed
print(shape.id)

# The element we are processing
print(model.by_guid(shape.guid))
print(shape.geometry.id)

#Extract placement and transformation matrix
matrix= shape.transformation.matrix
matrix= ifcopenshell.util.shape.get_shape_matrix(shape)
print(f"location: {matrix[:,3][0:3]}")

# Access vertices, edges, and faces:
vert= shape.geometry.verts
print(f"vertices: {vert}")

edges= shape.geometry.edges
print(f"edges: {edges}")

# faces always are triangles
faces= shape.geometry. faces
print(f"there are {len(faces)} faces: {faces}")

#%%
# 3. Quantities & Material Take-off (QTO)
import ifcopenshell.util.element

beam = model.by_type("IfcBeam")[0]
psets = ifcopenshell.util.element.get_psets(beam)
print(psets)  # Includes lengths, cross-section, etc.

# Get only properties and not quantities
print(ifcopenshell.util.element.get_psets(beam, psets_only=True))

# Get only quantities and not properties
print(ifcopenshell.util.element.get_psets(beam, qtos_only=True))

#%%
#4. Unit Conversion
import ifcopenshell.util.unit

unit_scale = ifcopenshell.util.unit.calculate_unit_scale(model)
print("Unit scale:", unit_scale)

# Find quantity set "BaseQuantities" for this 
ifc_project_length=None
for rel in beam.IsDefinedBy:
 if rel.is_a("IfcRelDefinesByProperties") and rel.RelatingPropertyDefinition.is_a("IfcElementQuantity"):
	 qto = rel.RelatingPropertyDefinition
	 if qto.Name == "BaseQuantities":
		 for quantity in qto.Quantities:
			 if quantity.is_a("IfcQuantityLength") and quantity.Name.lower() == "length":
				 ifc_project_length = quantity.LengthValue
				 break

# Check if we found a length value
if ifc_project_length is not None:
 # Convert to SI
 si_meters = ifc_project_length * unit_scale
 print(f"Original (project units): {ifc_project_length}")
 print(f"Converted to SI meters: {si_meters}")

 # Convert back to project units
 back_to_project_units = si_meters / unit_scale
 print(f"Back to project units: {back_to_project_units}")
else:
 print("No length quantity found in BaseQuantities for this element.")


#%%
#5. Custom IFC File Generation (copy and create a model)

## Copy an entity instance
import ifcopenshell.api.root
wall = model.by_type("IfcWall")[0]
wall_copy_class = ifcopenshell.api.root.copy_class(model, product = wall)

print(wall_copy_class)

import ifcopenshell.util.element

wall_shallow_copy = ifcopenshell.util.element.copy(model, wall)

print(wall_shallow_copy)

f = ifcopenshell.open("Infra-Bridge.ifc")
g = ifcopenshell.file(schema=f.schema)
beams = f.by_type("IfcBeam")
print(len(beams))
for b in beams[:1]:
	print(b)
	print ("__________________________")
	g.add(b)
g.write("copy_Beam_Infra-Bridge.ifc")






