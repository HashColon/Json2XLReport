import json
import os
import plotly.express as px
import pandas

imageSize = {
    "width": 960,
    "height": 540
}

# Create image of Silhouette graph on all data
# Silhouettes: 2D Array of silhouette data by cluster, item


def CreateImage_SilhouetteGraph_All(ClusterNames, Silhouettes_allClusters, filePath):
    # flatten silhouete_allClusters
    Flat_ClusterNames = []
    Flat_Silhouettes = []
    for i, SilhouetteGroup in enumerate(Silhouettes_allClusters):
        for silhouette in SilhouetteGroup:
            Flat_ClusterNames.append(ClusterNames[i])
            Flat_Silhouettes.append(silhouette)

    # Build dictionary from inputs
    Dict_GraphData = dict(
        [('ClusterName', Flat_ClusterNames), ('Silhouette', Flat_Silhouettes)])
    # build data frame for plotly graph
    DF_GraphData = pandas.DataFrame.from_dict(Dict_GraphData)
    # sort data frame by cluster number and silhouette value
    DF_GraphData.sort_values(
        by=['ClusterName', 'Silhouette'],
        ascending=True, inplace=True, ignore_index=True)
    # build Graphfigure
    GraphFig = px.bar(
        data_frame=DF_GraphData, width=imageSize['width'], height=imageSize['height'],
        x='Silhouette', color='ClusterName', range_x=[-1, 1])
    # export GraphFig
    GraphFig.write_image(
        filePath, width=imageSize['width'], height=imageSize['height'])
    return GraphFig


def CreateImage_SilhouetteGraph(Silhouettes, filePath):
    # Build dictionary from inputs
    Dict_GraphData = dict(
        [('Silhouette', Silhouettes)])
    # build data frame for plotly graph
    DF_GraphData = pandas.DataFrame.from_dict(Dict_GraphData)
    # sort data frame by cluster number and silhouette value
    DF_GraphData.sort_values(
        by=['Silhouette'],
        ascending=True, inplace=True, ignore_index=True)
    # build Graphfigure
    GraphFig = px.bar(
        data_frame=DF_GraphData, x='Silhouette',
        width=imageSize['width'], height=imageSize['height'], range_x=[-1, 1])
    # export GraphFig
    GraphFig.write_image(
        filePath, width=imageSize['width'], height=imageSize['height'])
    return GraphFig


def WriteEvaluationImages(ClusterNames, Silhouettes, ExpName, ExpRootDir):
    # Create Report image directory
    ReportImageDir = ExpRootDir + "/Report/Images"
    if not os.path.isdir(ReportImageDir):
        os.mkdir(ReportImageDir)

    # Create Silhouette graph for whole data
    SilhouetteGraphAllPath = ReportImageDir + "/" + ExpName + "_AllSilhouettes.jpg"
    allimg = CreateImage_SilhouetteGraph_All(
        ClusterNames, Silhouettes, SilhouetteGraphAllPath)

    # Create Silhouette graph for each clusters
    imgs = []
    for i in range(0, ClusterNames):
        SilhouetteGraphPath_i = ReportImageDir + "/" + \
            ExpName + "_Silhouettes_Cluster_" + str(i) + ".jpg"
        img = CreateImage_SilhouetteGraph(
            Silhouettes[i], SilhouetteGraphPath_i)
        imgs.append(img)

    return {'AllSilhouette': allimg, 'Silhouettes': imgs}
