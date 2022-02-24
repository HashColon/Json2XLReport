import FelineReportXL
import TrajectoryImage
import EvaluationImage
import json
import sys
import os
from datetime import datetime
import pickle


def NowString():
    return "[" + datetime.now().isoformat(" ", "milliseconds") + "]"


def Log(msg):
    print(NowString() + " FelineReport(Python): " + msg)


def TrajectoryImages(ExpRootDir, JsonData):
    ReportImagesDir = ExpRootDir + '/Report/Images'

    # Make Image: All input trajectories: InputTrajs
    Path_InputTrajs = ReportImagesDir + '/' + \
        JsonData['ExperimentName'] + '_InputTrajectories.jpg'
    GeoJsonList_InputTrajs = os.listdir(
        ExpRootDir + '/Inputs/GeoJson')
    TagList_InputTrajs = []
    for i, geojsonPath in enumerate(GeoJsonList_InputTrajs):
        GeoJsonList_InputTrajs[i] = ExpRootDir + \
            '/Inputs/GeoJson/' + geojsonPath
        TagList_InputTrajs.append("Input")
    TrajectoryImage.CreateImage_ManyGeoJsonFilePath(
        GeoJsonList_InputTrajs,
        TagList_InputTrajs,
        Path_InputTrajs)
    JsonData['Images']['InputTrajs'] = Path_InputTrajs

    # Make Image: Clsutered trajectories: ClusteringResult (with/wiithout noise)
    Path_ClusteringResult = ReportImagesDir + '/' + \
        JsonData['ExperimentName'] + '_ClusteringResult.jpg'
    Path_ClusteringResult_withNoise = ReportImagesDir + \
        '/ClusteringResult_withNoise.jpg'
    GeoJsonList_ClusteringResult_withNoise = []
    TagList_ClusteringResult_withNoise = []
    for i, clusterName in enumerate(JsonData['ClusterNames']):
        TagList_ClusteringResult_withNoise.append(
            JsonData["ExperimentName"] + "_" + clusterName + ".json")
        GeoJsonList_ClusteringResult_withNoise.append(
            ExpRootDir + "/Results/GeoJson/" + TagList_ClusteringResult_withNoise[i])
    if JsonData['IsNoiseClusterDefined']:
        GeoJsonList_ClusteringResult = GeoJsonList_ClusteringResult_withNoise[1:]
        TagList_ClusteringResult = TagList_ClusteringResult_withNoise[1:]
    else:
        GeoJsonList_ClusteringResult = GeoJsonList_ClusteringResult_withNoise
        TagList_ClusteringResult = TagList_ClusteringResult_withNoise

    TrajectoryImage.CreateImage_ManyGeoJsonFilePath(
        GeoJsonList_ClusteringResult,
        TagList_ClusteringResult,
        Path_ClusteringResult)
    JsonData['Images']['ClusteringResult'] = Path_ClusteringResult

    if JsonData['IsNoiseClusterDefined']:
        TrajectoryImage.CreateImage_ManyGeoJsonFilePath(
            GeoJsonList_ClusteringResult_withNoise,
            TagList_ClusteringResult_withNoise,
            Path_ClusteringResult_withNoise)
        JsonData['Images']['ClusteringResult_withNoise'] = Path_ClusteringResult_withNoise

    # Make Image: Clustered trajectories by each cluster: Clusters
    GeoJsonList_Representatives = []
    TagList_Representatives = []
    JsonData['Images']['Clusters'] = []
    for i, cName in enumerate(JsonData['ClusterNames']):
        filename = JsonData['ExperimentName'] + '_' + cName
        GeoJson_Cluster = ExpRootDir + '/Results/GeoJson/' + filename + '.json'
        repfilename = JsonData['ExperimentName'] + '_Representative_' + cName
        GeoJson_Rep = ExpRootDir + '/Results/GeoJson/' + repfilename + '.json'
        Path_Cluster = ExpRootDir + '/Results/Images/' + filename + '.jpg'

        if cName != "Noise":
            GeojsonList = [GeoJson_Cluster, GeoJson_Rep]
            TagList = [filename, repfilename]
            # add representative infos
            GeoJsonList_Representatives.append(GeoJson_Rep)
            TagList_Representatives.append(repfilename)
        else:
            GeojsonList = [GeoJson_Cluster]
            TagList = [filename]
        TrajectoryImage.CreateImage_ManyGeoJsonFilePath(
            GeojsonList, TagList, Path_Cluster, showtags=False)
        JsonData['Images']['Clusters'].append(Path_Cluster)

    # Make Image: Representative traajectories
    Path_Representative = ReportImagesDir + '/' + \
        JsonData['ExperimentName'] + '_Representative.jpg'
    TrajectoryImage.CreateImage_ManyGeoJsonFilePath(
        GeoJsonList_Representatives, TagList_Representatives,
        Path_Representative)
    JsonData['Images']['Representatives'] = Path_Representative


def EvaluationImages(ExpRootDir, JsonData):
    ExpName = JsonData['ExperimentName']
    ClusterNames = JsonData["ClusterNames"]
    Silhouettes = JsonData['Evaluation']['Silhouettes']
    # Silhouette all
    Path_SilhouetteGraphAll = ExpRootDir + \
        "/Report/Images/" + ExpName + "_SilhouettesAll.jpg"
    EvaluationImage.CreateImage_SilhouetteGraph_All(
        ClusterNames, Silhouettes, Path_SilhouetteGraphAll)
    JsonData['Images']['Silhouettes'] = Path_SilhouetteGraphAll

    # Silhouettes by each cluseter
    JsonData['Images']['SilhouettesCluster'] = []
    for i, clusterName in enumerate(ClusterNames):
        Path_SilhouetteGraph = ExpRootDir + "/Report/Images/" + \
            ExpName + "_Silhouettes" + clusterName + ".jpg"
        EvaluationImage.CreateImage_SilhouetteGraph(
            Silhouettes[i], Path_SilhouetteGraph)
        JsonData['Images']['SilhouettesCluster'].append(Path_SilhouetteGraph)


def NoiseClusterModification(jsonData):
    if jsonData['IsNoiseClusterDefined']:
        if ("Noise" in jsonData['ClusterNames']):
            noiseIdx = jsonData['ClusterNames'].index("Noise")
            # if noise cluster is empty, than erase noise cluster
            if not jsonData["ClusteringResults"][noiseIdx]:
                # IsNoiseClusterDefined
                jsonData['IsNoiseClusterDefined'] = False
                # ClusterNames
                jsonData['ClusterNames'].pop(0)
                # ClusteringResults
                jsonData['ClusteringResults'].pop(0)
                # Evaluation.DaviesBouldin
                jsonData['Evaluation']['DaviesBouldin'].pop(0)
                # Silhouettes
                jsonData['Evaluation']['Silhouettes'].pop(0)


def PrepFelineReport(ExpRootDir):
    # check if ExpRootDir exists
    if not os.path.isdir(ExpRootDir):
        Log("Invalid Feline result root directory: " + ExpRootDir)
        exit()
    else:
        ExpName = ExpRootDir.split("/")[-1]

    # Start the shits
    Log("Creating reports for Feline clustering results.")

    # Read reports in json
    JsonFile_AppReport_Path = ExpRootDir + "/Report/" + ExpName + "_Report.json"
    JsonFile_AppReport = open(JsonFile_AppReport_Path, "r")
    Json_AppReport = json.load(JsonFile_AppReport)

    # build image directories
    ReportImagesDir = ExpRootDir + '/Report/Images'
    os.makedirs(ReportImagesDir, exist_ok=True)
    ResultsImagesDir = ExpRootDir + '/Results/Images'
    os.makedirs(ResultsImagesDir, exist_ok=True)
    Log("Finished creating directories for export files.")

    # add image obj to json data
    Json_AppReport['Images'] = {}

    # check if noise cluster is empty
    NoiseClusterModification(Json_AppReport)

    # create evaluation graph images
    EvaluationImages(ExpRootDir, Json_AppReport)
    Log("Finished exporting evaluation images.")

    # create Result images
    TrajectoryImages(ExpRootDir, Json_AppReport)
    Log("Finished exporting trajectory images.")

    return Json_AppReport


if __name__ == "__main__":

    for i, arg in enumerate(sys.argv[1:-2]):
        ExpRootDir = arg
        JsonReport = PrepFelineReport(ExpRootDir)

        # # pickling for test
        # with open('testpickle.pkl', 'wb') as f:
        #     pickle.dump(JsonReport, f)

        # loading pickle for test
        # with open('testpickle.pkl', 'rb') as f:
        #     JsonReport = pickle.load(f)

        # print JsonReport at screen
        # checkingReportResult = json.dumps(JsonReport, indent=2)
        # Log("Report:")
        # print(checkingReportResult)
        FelineReportXL.FelineReportXL(JsonReport, ExpRootDir)
        Log("Finished exporting report excel file.")

        if len(sys.argv) > 1:
            RemoteCopyDir = sys.argv[-2]
            Log("Copying to remote destination." + RemoteCopyDir)
            PassStr = sys.argv[-1]
            ScpBashStr = 'sshpass -p "' + PassStr + '" scp -r ' + \
                ExpRootDir + " " + RemoteCopyDir
            os.system(ScpBashStr)
        Log("Creating feline report finished. (" + str(i) + ')')
