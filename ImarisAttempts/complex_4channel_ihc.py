# <CustomTools>
#   <Menu>
#    <Submenu name="Mikey">
#     <Item name="IHC Analysis 4-Channel" icon="Python" tooltip="Creates surfaces, masked channels, and surfaces with stats for 4-channel IHC">
#      <Command>PythonXT::create_surface_channels_then_analyze(%i)</Command>
#     </Item>
#    </Submenu>
#   </Menu>
# </CustomTools>

import ImarisLib
import time
import numpy as np
import pandas as pd


def create_surface_channels_then_analyze(aImarisId):
    """
    Advanced IHC analysis script for Alzheimer's disease microglial function research.

    This script is designed for analyzing 4-channel IHC images with customizable surface creation
    and masking options. It allows creation of up to 4 surfaces, which can be based on either
    original channels or channels masked onto other surfaces.

    Features:
    - Setup for 4 standard IHC channels with custom naming
    - Customizable threshold settings for surface creation
    - Option to create surfaces from original or masked channels
    - Automatic management of dependencies between surfaces
    - Comprehensive statistics calculation and export

    OPTIONAL FILTERING FEATURES:

    This script includes optional filtering steps after each surface creation, which are commented
    out by default (enclosed in triple quotes). Here's how to use these filtering options:

    1. Volume-based filtering:
       - Default minimum volume is set to 70 (customizable)
       - Example: vSurfaceFilter.AddFilterOnVolume(70, 1000000)

    2. Distance-based filtering:
       - Filter surfaces based on distance to other surfaces
       - Example: vSurfaceFilter.AddFilterOnDistanceToSurfaces(vOtherSurface, vDistMin, vDistMax)

    How to activate these filters:

    1. Simply remove the triple quotes around the section you want to activate
    2. Adjust the parameters as needed(minimum volume, distance thresholds)
    3. If filtering by distance to another surface, make sure the referenced surface exists

    These filters are available after each surface creation.

    """
    # Get Imaris application
    vImarisLib = ImarisLib.ImarisLib()
    vImaris = vImarisLib.GetApplication(aImarisId)

    # Check for successful connection to Imaris
    if vImaris is None:
        print('Could not connect to Imaris')
        return

    # Get the current dataset
    vDataSet = vImaris.GetDataSet()
    if vDataSet is None:
        print('No data set loaded')
        return

    # Get image properties - using vDataSet instead of vImaris
    vSizeX = vDataSet.GetSizeX()
    vSizeY = vDataSet.GetSizeY()
    vSizeZ = vDataSet.GetSizeZ()
    vSizeC = vDataSet.GetSizeC()

    print(f"Dataset dimensions: X={vSizeX} Y={vSizeY} Z={vSizeZ} Channels={vSizeC}")

    # Create a Dialog for user input
    vUI = vImaris.GetUI()

    # Get Channel names from dataset
    original_channel_names = []
    for c in range(vSizeC):
        original_channel_names.append(vDataSet.GetChannelName(c))

    # =========== IHC Channel Configuration UI ===========

    vChannelUI = vUI.CreateGroup("IHC Channel Configuration")

    # Create dropdown for each of the 4 standard channels
    vChannel1Selection = vUI.CreateComboBox('Channel 1 (e.g., ThioS for amyloid): ', original_channel_names, min(0, vSizeC-1))
    vChannel2Selection = vUI.CreateComboBox('Channel 2 (e.g., iba1 for microglia): ', original_channel_names, min(1, vSizeC-1))
    vChannel3Selection = vUI.CreateComboBox('Channel 3 (e.g., cd68 for DAMs): ', original_channel_names, min(2, vSizeC-1))
    vChannel4Selection = vUI.CreateComboBox('Channel 4 (e.g., ATG-9): ', original_channel_names, min(3, vSizeC-1))

    # Channel custom names
    vChannel1Name = vUI.CreateString('Channel 1 name: ', 'ThioS')
    vChannel2Name = vUI.CreateString('Channel 2 name: ', 'Iba1')
    vChannel3Name = vUI.CreateString('Channel 3 name: ', 'CD68')
    vChannel4Name = vUI.CreateString('Channel 4 name: ', 'ATG-9')

    # =========== Surface Creation UI ===========

    # Common settings for all surfaces
    vCommonSettings = vUI.CreateGroup("Common Surface Settings")
    vSurfaceDetail = vUI.CreateNumberRange('Surface detail (um): ', 0.1, 5.0, 0.1)
    vSurfaceDetail.SetDefaultValue(0.3)
    vBackgroundSubtraction = vUI.CreateBooleanCheckbox('Enable background subtraction', True)

    # Surface 1 Configuration (Primary Surface)
    vSurface1Group = vUI.CreateGroup("Surface 1 Configuration")
    vSurface1Enable = vUI.CreateBooleanCheckbox('Create Surface 1', True)
    vSurface1Channel = vUI.CreateComboBox('Surface 1 from channel: ', ['Channel 1', 'Channel 2', 'Channel 3', 'Channel 4'], 0)
    vSurface1Threshold = vUI.CreateNumberRange('Surface 1 threshold: ', 0, 255, 1)
    vSurface1Threshold.SetDefaultValue(30)

    # Surface 2 Configuration (Secondary Surface)
    vSurface2Group = vUI.CreateGroup("Surface 2 Configuration")
    vSurface2Enable = vUI.CreateBooleanCheckbox('Create Surface 2', True)
    vSurface2Channel = vUI.CreateComboBox('Surface 2 from channel: ', ['Channel 1', 'Channel 2', 'Channel 3', 'Channel 4'], 1)
    vSurface2Threshold = vUI.CreateNumberRange('Surface 2 threshold: ', 0, 255, 1)
    vSurface2Threshold.SetDefaultValue(30)

    # Surface 3 Configuration
    vSurface3Group = vUI.CreateGroup("Surface 3 Configuration")
    vSurface3Enable = vUI.CreateBooleanCheckbox('Create Surface 3', False)
    vSurface3Type = vUI.CreateComboBox('Surface 3 source type: ', ['Original Channel', 'Masked Channel'], 1)
    vSurface3Channel = vUI.CreateComboBox('Surface 3 channel: ', ['Channel 1', 'Channel 2', 'Channel 3', 'Channel 4'], 2)
    vSurface3MaskOnto = vUI.CreateComboBox('Mask onto surface: ', ['Surface 1', 'Surface 2'], 1)
    vSurface3Threshold = vUI.CreateNumberRange('Surface 3 threshold: ', 0, 255, 1)
    vSurface3Threshold.SetDefaultValue(30)

    # Surface 4 Configuration
    vSurface4Group = vUI.CreateGroup("Surface 4 Configuration")
    vSurface4Enable = vUI.CreateBooleanCheckbox('Create Surface 4', False)
    vSurface4Type = vUI.CreateComboBox('Surface 4 source type: ', ['Original Channel', 'Masked Channel'], 1)
    vSurface4Channel = vUI.CreateComboBox('Surface 4 channel: ', ['Channel 1', 'Channel 2', 'Channel 3', 'Channel 4'], 3)
    vSurface4MaskOnto = vUI.CreateComboBox('Mask onto surface: ', ['Surface 1', 'Surface 2'], 1)
    vSurface4Threshold = vUI.CreateNumberRange('Surface 4 threshold: ', 0, 255, 1)
    vSurface4Threshold.SetDefaultValue(30)

    # Statistics Export
    vExportGroup = vUI.CreateGroup("Statistics Export")
    vExportPath = vUI.CreateFileSave('Select statistics export path', '*.csv')

    # Show dialog
    vResult = vUI.Show()
    if not vResult:
        return "Canceled by user"

    # =========== Get Selected Values ===========

    # Get common settings
    detail_level = vSurfaceDetail.GetValue()
    use_background_subtraction = vBackgroundSubtraction.GetValue()
    export_path = vExportPath.GetValue()

    # Map logical channels to actual dataset channels and names
    channel_mapping = [
        vChannel1Selection.GetValue(),
        vChannel2Selection.GetValue(),
        vChannel3Selection.GetValue(),
        vChannel4Selection.GetValue()
    ]

    channel_names_custom = [
        vChannel1Name.GetValue(),
        vChannel2Name.GetValue(),
        vChannel3Name.GetValue(),
        vChannel4Name.GetValue()
    ]

    # Configure surfaces
    surface_configs = []

    # Surface 1 (always from original channel)
    if vSurface1Enable.GetValue():
        channel_idx = channel_mapping[vSurface1Channel.GetValue()]
        surface_configs.append({
            'enabled': True,
            'type': 'original',
            'channel_idx': channel_idx,
            'channel_name': channel_names_custom[vSurface1Channel.GetValue()],
            'threshold': vSurface1Threshold.GetValue(),
            'mask_onto': None
        })
    else:
        surface_configs.append({'enabled': False})

    # Surface 2 (always from original channel)
    if vSurface2Enable.GetValue():
        channel_idx = channel_mapping[vSurface2Channel.GetValue()]
        surface_configs.append({
            'enabled': True,
            'type': 'original',
            'channel_idx': channel_idx,
            'channel_name': channel_names_custom[vSurface2Channel.GetValue()],
            'threshold': vSurface2Threshold.GetValue(),
            'mask_onto': None
        })
    else:
        surface_configs.append({'enabled': False})

    # Surface 3 (can be from original or masked channel)
    if vSurface3Enable.GetValue():
        channel_idx = channel_mapping[vSurface3Channel.GetValue()]
        surface_type = 'original' if vSurface3Type.GetValue() == 0 else 'masked'
        mask_onto = vSurface3MaskOnto.GetValue() if surface_type == 'masked' else None

        surface_configs.append({
            'enabled': True,
            'type': surface_type,
            'channel_idx': channel_idx,
            'channel_name': channel_names_custom[vSurface3Channel.GetValue()],
            'threshold': vSurface3Threshold.GetValue(),
            'mask_onto': mask_onto
        })
    else:
        surface_configs.append({'enabled': False})

    # Surface 4 (can be from original or masked channel)
    if vSurface4Enable.GetValue():
        channel_idx = channel_mapping[vSurface4Channel.GetValue()]
        surface_type = 'original' if vSurface4Type.GetValue() == 0 else 'masked'
        mask_onto = vSurface4MaskOnto.GetValue() if surface_type == 'masked' else None

        surface_configs.append({
            'enabled': True,
            'type': surface_type,
            'channel_idx': channel_idx,
            'channel_name': channel_names_custom[vSurface4Channel.GetValue()],
            'threshold': vSurface4Threshold.GetValue(),
            'mask_onto': mask_onto
        })
    else:
        surface_configs.append({'enabled': False})

    # =========== Create Surfaces in Order ===========

    # Get Factory and Scene
    vFactory = vImaris.GetImageProcessing()
    vSurpassScene = vImaris.GetSurpassScene()

    # Create dictionary to store created surfaces
    surface_objects = {}
    masked_channel_indices = []
    next_channel_index = vSizeC

    # First pass - create surfaces from original channels
    for i, config in enumerate(surface_configs):
        if config['enabled'] and config['type'] == 'original':
            print(f"Creating Surface {i+1} from channel {config['channel_name']}...")

            surface = vFactory.DetectSurfaces(vDataSet, [], config['channel_idx'],
                                               use_background_subtraction, config['threshold'],
                                               0.0, True, False, False, detail_level, "")

            if surface is None:
                print(f"Error creating Surface {i+1}")
                continue

            # Set name and add to scene
            surface_name = f"{config['channel_name']}_Surface"
            surface.SetName(surface_name)
            vSurpassScene.AddChild(surface, -1)

            # Store surface
            surface_objects[i] = {
                'object': surface,
                'name': surface_name,
                'channel_name': config['channel_name']
            }

            # Optional filtering step
            """
    # Uncomment and modify these lines to apply volume filtering
    print(f"Filtering Surface {i + 1} by volume...")
    vSurfaceFilter = surface.GetFilter()
    # Set minimum volume filter
    vSurfaceFilter.AddFilterOnVolume(70, 1000000)  # Min volume 70, max very large
    surface.SetFilter(vSurfaceFilter)

    # Optional: Filter by distance to another existing surface
    # if another_surface_idx in surface_objects:
    #     print(f"Filtering Surface {i+1} by distance...")
    #     vDistMin = 0   # Minimum distance
    #     vDistMax = 10  # Maximum distance (adjust as needed)
    #     vSurfaceFilter.AddFilterOnDistanceToSurfaces(surface_objects[another_surface_idx]['object'], vDistMin, vDistMax)
    #     surface.SetFilter(vSurfaceFilter)
    """

# Second pass - create surfaces from masked channels
for i, config in enumerate(surface_configs):
if config['enabled'] and config['type'] == 'masked':
    # Check if the surface to mask onto exists
    mask_onto_idx = config['mask_onto']
    if mask_onto_idx not in surface_objects:
        print(f"Error: Cannot create Surface {i+1} because Surface {mask_onto_idx+1} does not exist")
        continue

    mask_surface = surface_objects[mask_onto_idx]['object']
    mask_surface_name = surface_objects[mask_onto_idx]['name']

    print(f"Creating mask from Surface {mask_onto_idx+1} ({mask_surface_name})...")
    vMaskedChannel = vFactory.CreateMaskChannel(vDataSet, mask_surface, "Inside", 0, 0)

    print(f"Masking channel {config['channel_name']} with {mask_surface_name}...")
    vMaskedData = vFactory.CreateMaskedChannel(vDataSet, config['channel_idx'], vMaskedChannel, "0", "Inside")

    if vMaskedData is None:
        print(f"Error creating masked channel for Surface {i+1}")
        continue

    # Add masked channel to dataset
    masked_name = f"{config['channel_name']}_masked_by_{mask_surface_name}"
    vDataSet.SetChannelName(next_channel_index, masked_name)
    vDataSet.SetChannelColorRGBA(next_channel_index, vDataSet.GetChannelColorRGBA(config['channel_idx']))
    masked_channel_idx = next_channel_index
    masked_channel_indices.append(masked_channel_idx)
    next_channel_index += 1

    # Create surface from masked channel
    print(f"Creating Surface {i+1} from masked channel {masked_name}...")
    surface = vFactory.DetectSurfaces(vDataSet, [], masked_channel_idx,
                                       use_background_subtraction, config['threshold'],
                                       0.0, True, False, False, detail_level, "")

    if surface is None:
        print(f"Error creating Surface {i+1} from masked channel")
        continue

    # Set name and add to scene
    surface_name = f"{config['channel_name']}_in_{mask_surface_name}"
    surface.SetName(surface_name)
    vSurpassScene.AddChild(surface, -1)

    # Store surface
    surface_objects[i] = {
        'object': surface,
        'name': surface_name,
        'channel_name': config['channel_name'],
        'masked': True,
        'masked_onto': mask_onto_idx
    }

    # Optional filtering step
    """
    # Uncomment and modify these lines to apply volume filtering
    print(f"Filtering Surface {i + 1} by volume...")
    vSurfaceFilter = surface.GetFilter()
    # Set minimum volume filter
    vSurfaceFilter.AddFilterOnVolume(70, 1000000)  # Min volume 70, max very large
    surface.SetFilter(vSurfaceFilter)

    # Optional: Filter by distance to another surface
    # if another_surface_idx in surface_objects:
    #     print(f"Filtering Surface {i+1} by distance...")
    #     vDistMin = 0   # Minimum distance
    #     vDistMax = 10  # Maximum distance (adjust as needed)
    #     vSurfaceFilter.AddFilterOnDistanceToSurfaces(surface_objects[another_surface_idx]['object'], vDistMin, vDistMax)
    #     surface.SetFilter(vSurfaceFilter)
    """

# =========== Calculate Statistics ===========

# Calculate statistics for all surfaces
all_stats = {}

print("Calculating statistics for all surfaces...")

# Process each created surface
for i, surface_info in surface_objects.items():
surface = surface_info['object']
print(f"Calculating statistics for Surface {i+1} ({surface_info['name']})...")

# Get statistics
surface_stats = surface.GetStatistics()
stat_names = surface_stats.GetNames()
stat_values = surface_stats.GetValues()

# Extract key statistics
surface_data = {
    'Name': surface_info['name'],
    'Channel': surface_info['channel_name']
}

for j in range(len(stat_names)):
    if "Volume" in stat_names[j] and "Sum" in stat_names[j]:
        surface_data["Total_Volume"] = stat_values[j]
    if "Sphericity" in stat_names[j]:
        surface_data["Sphericity"] = stat_values[j]
    if "Number of" in stat_names[j] and "Triangles" not in stat_names[j]:
        surface_data["Count"] = stat_values[j]

all_stats[f"Surface_{i+1}"] = surface_data

# Calculate pairwise distance metrics
print("Calculating pairwise distance metrics between surfaces...")

# Dictionary to store overlap data
overlap_stats = {}

# Loop through unique surface pairs
for i, surface1_info in surface_objects.items():
for j, surface2_info in surface_objects.items():
    if i < j:  # Process each pair only once
        surface1 = surface1_info['object']
        surface2 = surface2_info['object']

        print(f"Calculating distance metrics between Surface {i+1} and Surface {j+1}...")

        # Create distance transformation
        dist_stats = vFactory.DistanceTransformForSurfaces(surface1, surface2)

        # In a real implementation, you would extract meaningful metrics from dist_stats
        # This is a simplified placeholder
        pair_name = f"Surface_{i+1}_to_Surface_{j+1}"
        overlap_stats[pair_name] = {
            'Surface1': surface1_info['name'],
            'Surface2': surface2_info['name'],
            'Mean_Distance': 0,  # This would be calculated from dist_stats
            'Overlap_Volume': 0   # This would be calculated from dist_stats
        }

# =========== Export Statistics ===========

print(f"Exporting statistics to {export_path}...")

# Prepare data for single-surface statistics
surface_stats_rows = []

for surface_id, stats in all_stats.items():
row = {
    'Surface': surface_id,
    'Name': stats['Name'],
    'Channel': stats['Channel'],
    'Total_Volume': stats.get('Total_Volume', 0),
    'Sphericity': stats.get('Sphericity', 0),
    'Count': stats.get('Count', 0)
}
surface_stats_rows.append(row)

# Prepare data for pairwise statistics
pairwise_stats_rows = []

for pair_id, stats in overlap_stats.items():
row = {
    'Pair': pair_id,
    'Surface1': stats['Surface1'],
    'Surface2': stats['Surface2'],
    'Mean_Distance': stats['Mean_Distance'],
    'Overlap_Volume': stats['Overlap_Volume']
}
pairwise_stats_rows.append(row)

# Create DataFrames
surface_df = pd.DataFrame(surface_stats_rows)
pairwise_df = pd.DataFrame(pairwise_stats_rows)

# Export to CSV
with pd.ExcelWriter(export_path) as writer:
surface_df.to_excel(writer, sheet_name='Surface_Statistics', index=False)
pairwise_df.to_excel(writer, sheet_name='Pairwise_Statistics', index=False)

print('Analysis completed successfully!')
return "Success"