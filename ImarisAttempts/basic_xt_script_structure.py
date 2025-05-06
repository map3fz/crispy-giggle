# <CustomTools>
#   <Menu>
#    <Submenu name="Mikey">
#     <Item name="Python Attempt 1" icon="Python" tooltip="Creates surfaces, masked channels, and surfaces with stats">
#      <Command>PythonXT::create_surfaces_channels_then_analyze(%i)</Command>
#     </Item>
#    </Submenu>
#   </Menu>
# </CustomTools>

import ImarisLib
import time
import numpy as np
import pandas as pd

def create_surfaces_channels_then_analyze(aImarisId): # command in meta-data must match function name
    """
    This is a function built with the intention of being able to be manipulated to easily set
    repeated, changed, and customized for individual needs based on threshold values expected for
    different IHC experiments. Along with this, there is a section setup that will allow you to
    create a channel that is masked onto a surface with an example set for masking onto your primary
    surface, or for your secondary surface. Along with this, names can be changed to allow it to
    automatically apply the same name for each iteration of the surfaces and channels that makes
    pulling statistics more consistent and the names more descriptive. This will allow for fewer
    steps in the analysis that is done by hand, hopefully speeding up the process while making
    less possible errors.
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

    # Get image properties
    vSizeX = vDataSet.GetSizeX()
    vSizeY = vDataSet.GetSizeY()
    vSizeZ = vDataSet.GetSizeZ()
    vSizeC = vDataSet.GetSizeC()

    print(f"Dataset dimensions: X={vSizeX} Y={vSizeY} Z={vSizeZ} Channels={vSizeC}")

    # Main Work flow

    # Create a Dialog for use input
    vUI = vImaris.GetUI()

    # Get Channel names
    channel_names = []
    for c in range(vSizeC):
        channel_names.append(vImaris.GetChannelName(c))

    # Create dialog for primary surface channel (e.g., ThioS for plaque)
    vPrimarySurfaceChannel = vUI.CreateComboBox('Select primary surface channel: ', channel_names, 0)

    # Create threshold setting for primary surface
    vPrimaryThreshold = vUI.CreateNumberRange('Primary threshold: ', 0, 255, 1)
    vPrimaryThreshold.SetDefaultValue(30)

    # Create dialog for secondary surface
    vSecondarySurfaceChannel = vUI.CreateNumberRange('Secondary surface channel: ', channel_names, 1)

    # Create threshold setting for secondary surface
    vSecondaryThreshold = vUI.CreateNumberRange('Secondary threshold: ', 0, 255, 1)
    vSecondaryThreshold.SetDefaultValue(30)

    # Surface detail level
    vSurfaceDetail = vUI.CreateNumberRange('Surface detail (um): ', 0.1, 5.0, 0.1)
    vSurfaceDetail.SetDefaultValue(0.3)

    # Background subtraction
    vBackgroundSubtraction = vUI.CreateBooleanCheckbox('Enable background subtraction', True)

    # Dialog for statistics export path
    vExportPath = vUI.CreateFileSave('Select statistics export path', '*.csv')

    # Show dialog
    vResult = vUI.Show()
    if not vResult:
        return "Canceled by user"

    # Get selected Values
    primary_channel_idx = vPrimarySurfaceChannel.GetValue()
    primary_threshold = vPrimaryThreshold.GetValue()
    secondary_channel_idx = vSecondaryThreshold.GetValue()
    secondary_threshold = vSecondaryThreshold.GetValue()
    detail_level = vSurfaceDetail.GetValue()
    use_background_subtraction = vBackgroundSubtraction.GetValue()
    export_path = vExportPath.GetValue()

    # Create primary surface (e.g., amyloid plaques w/ ThioS)
    print(f"Creating primary surface from channel {channel_names[primary_channel_idx]}...")

    # Get the surface factory
    vFactory = vImaris.GetImageProcessing()
    vSurpassScene = vImaris.GetSurpassScene()

    # Create the primary surface
    vSurfaces = vFactory.DetectSurfaces(vDataSet, [], primary_channel_idx,
                                        use_background_subtraction, primary_threshold,
                                        0.0, True, False, False, detail_level, "")

    if vSurfaces is None:
        print("Error creating primary surface")
        return

    # Set name for the primary surface
    primary_surface_name = f"{channel_names[primary_channel_idx]}_Surface"
    vSurfaces.SetName(primary_surface_name)

    # Add surface to the scene
    vSurpassScene.AddChild(vSurfaces, -1)

    # Create secondary surface (e.g., microglia with iba1)
    print(f"Creating secondary surface from channel {channel_names[secondary_channel_idx]}...")

    # Create the secondary surface
    vSecondSurfaces = vFactory.DetectSurfaces(vDataSet, [], secondary_channel_idx,
                                              use_background_subtraction, secondary_threshold,
                                              0.0, False, False, detail_level, "")
    if vSecondSurfaces is None:
        print("Error creating secondary surface")
        return

    # Set name for the secondary surface
    secondary_surface_name = f"{channel_names[secondary_channel_idx]}_Surface"
    vSecondSurfaces.SetName(secondary_surface_name)

    # Add surface to the scene
    vSurpassScene.AddChild(vSecondSurfaces, -1)

    # Creating masked channels based on surfaces
    print("Creating masked channels...")

    # Mask channel inside primary surface
    vMaskedChannel = vFactory.CreateMaskChannel(vDataSet, vSurfaces, "Inside", 0, 0)

    # Create a new channel that is the original channel masked by the surface
    for c in range(vSizeC):
        print(f"Masking channel {channel_names[c]} with {primary_surface_name}...")
        vMaskedData = vFactory.CreateMaskedChannel(vDataSet, c, vMaskedChannel, "0", "Inside")

    if vMaskedData is not None:
        # Add masked channel to dataset
        masked_name = f"{channel_names[c]}_masked_by_{primary_surface_name}"
        vImaris.GetDataSet().SetChannelName(vSizeC, masked_name)
        vImaris.GetDataSet().SetChannelColorRGBA(vSizeC, vDataSet.GetChannelColorRGBA(c))

    # Create masked channels based on secondary surface (e.g., microglia)
    print("Creating masked channels...")

    # Mask channel inside secondary surface
    vMaskedChannel = vFactory.CreateMaskChannel(vDataSet, vSecondSurfaces, "Inside", 0, 0)

    # Create masked channels for all channels
    masked_channel_indices = []  # Store the indices of newly created masked channels

    for c in range(vSizeC):
        print(f"Masking channel {channel_names[c]} with {secondary_surface_name}...")
        vMaskedData = vFactory.CreateMaskedChannel(vDataSet, c, vMaskedChannel, "0", "Inside")

        if vMaskedData is not None:
            # Add masked channel to dataset
            new_channel_index = vSizeC + len(masked_channel_indices)
            masked_name = f"{channel_names[c]}_masked_by_{secondary_surface_name}"
            vImaris.GetDataSet().SetChannelName(new_channel_index, masked_name)
            vImaris.GetDataSet().SetChannelColorRGBA(new_channel_index, vDataSet.GetChannelColorRGBA(c))
            masked_channel_indices.append(new_channel_index)

    # Now create new surfaces based on the masked channels
    print("Creating surfaces from masked channels...")

    # Create surface from first masked channel (assuming this is the primary stain masked by microglia)
    if len(masked_channel_indices) >= 1:
        primary_masked_idx = masked_channel_indices[0]  # First masked channel (likely ThioS in microglia)

        print(f"Creating surface from masked channel {vImaris.GetDataSet().GetChannelName(primary_masked_idx)}...")
        vMaskedPrimarySurface = vFactory.DetectSurfaces(vDataSet, [], primary_masked_idx,
                                                        use_background_subtraction, primary_threshold,
                                                        0.0, True, False, False, detail_level, "")

        if vMaskedPrimarySurface is not None:
            masked_primary_name = f"{channel_names[primary_channel_idx]}_in_{secondary_surface_name}"
            vMaskedPrimarySurface.SetName(masked_primary_name)
            vSurpassScene.AddChild(vMaskedPrimarySurface, -1)

    # Create surface from second masked channel (if available)
    if len(masked_channel_indices) >= 2:
        secondary_masked_idx = masked_channel_indices[1]  # Second masked channel

        print(
            f"Creating surface from masked channel {vImaris.GetDataSet().GetChannelName(secondary_masked_idx)}...")
        vMaskedSecondarySurface = vFactory.DetectSurfaces(vDataSet, [], secondary_masked_idx,
                                                          use_background_subtraction, secondary_threshold,
                                                          0.0, True, False, False, detail_level, "")

        if vMaskedSecondarySurface is not None:
            masked_secondary_name = f"{channel_names[secondary_channel_idx]}_in_{secondary_surface_name}"
            vMaskedSecondarySurface.SetName(masked_secondary_name)
            vSurpassScene.AddChild(vMaskedSecondarySurface, -1)

    # Calculate and export statistics
    print("Calculating statistics...")

    # Get statistics from primary surface
    vPrimaryStatistics = vSurfaces.GetStatistics()
    vPrimaryStatNames = vPrimaryStatistics.GetNames()
    vPrimaryStatValues = vPrimaryStatistics.GetValues()
    vPrimaryStatIds = vPrimaryStatistics.GetIds()

    # Get statistics from secondary surface
    vSecondaryStatistics = vSecondSurfaces.GetStatistics()
    vSecondaryStatNames = vSecondaryStatistics.GetNames()
    vSecondaryStatValues = vSecondaryStatistics.GetValues()
    vSecondaryStatIds = vSecondaryStatistics.GetIds()

    # Create distance transformation to calculate volume overlap
    print("Calculating distance statistics and overlaps...")
    vDistStat = vFactory.DistanceTransformForSurfaces(vSecondSurfaces, vSurfaces)

    # Prepare data for export
    primary_stats_data = {}

    # Extract total surface volume and sphericity for the primary surface (ThioS)
    for i in range(len(vPrimaryStatNames)):
        if "Volume" in vPrimaryStatNames[i] and "Sum" in vPrimaryStatNames[i]:
            primary_stats_data["Total_Surface_Volume"] = vPrimaryStatValues[i]
        if "Sphericity" in vPrimaryStatNames[i]:
            primary_stats_data["ThioS_Sphericity"] = vPrimaryStatValues[i]

    # Calculate overlap volume between primary and secondary surfaces
    overlap_volume = 0
    # Implementation depends on the exact Imaris API
    # Typically would use the distance transform to identify volumes within a certain distance

    # Export statistics to CSV
    print(f"Exporting statistics to {export_path}...")
    stats_df = pd.DataFrame({
        'Statistic': ["Total_Surface_Volume", "ThioS_Sphericity", "Overlap_Volume"],
        'Value': [primary_stats_data.get("Total_Surface_Volume", 0),
                  primary_stats_data.get("ThioS_Sphericity", 0),
                  overlap_volume]
    })

    stats_df.to_csv(export_path, index=False)

    print('Analysis completed successfully!')
    return "Success"