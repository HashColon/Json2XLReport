import json
import os
import sys
from datetime import datetime
import EvaluationImage


def NowString():
    return "[" + datetime.now().isoformat(" ", "milliseconds") + "]"


def Log(msg):
    print(NowString() + " FelineReport(Python): " + msg)


def GetReport(ExpRootDir):
    # check if ExpRootDir exists
    if not os.path.isdir(ExpRootDir):
        Log("Invalid Feline result root directory: " + ExpRootDir)
        exit()
    else:
        ExpName = ExpRootDir.split("/")[-1]

    # Start the shits
    Log("Reading reports:" + ExpName)

    # Read reports in json
    JsonFile_Report_Path = ExpRootDir + "/Report/" + ExpName + "_Report.json"
    JsonFile_Report = open(JsonFile_Report_Path, "r")
    Json_Report = json.load(JsonFile_Report)

    return Json_Report


def GetMedianIndices(DistMatrix, Labels):
    l = len(Labels)
    median = [0] * l
    for c in range(0, l):
        minsum = -1
        for i in Labels[c]:
            sum_i = 0
            for j in Labels[c]:
                sum_i += DistMatrix[i][j]
            if (minsum < 0) or (minsum > sum_i):
                minsum = sum_i
                median[c] = i
    return median


def GetLabels(clsuterResults, N):
    labels = [0] * N
    checker = [0] * len(clusterResults)
    for i, items in enumerate(clsuterResults):
        for item in items:
            labels[item] = i
            checker[i] += 1
    checker2 = [0] * len(clusterResults)
    for i in range(len(clusterResults)):
        checker2[i] = len(clusterResults[i])

    checker3 = [0] * len(clusterResults)
    for i in labels:
        checker3[i] += 1

    return labels


def Silhouette_cluster(DistMatrix, clusterResults, opt1=True, optval=1.380, opt2=False):
    N = len(DistMatrix)
    l = len(clusterResults)
    labels = GetLabels(clusterResults, N)
    res = [[] for i in range(0, l)]

    for i in range(N):
        c = labels[i]
        sims = [0] * l
        a_i = 0
        b_i = -1
        for j in range(N):
            if i == j:
                continue
            c_j = labels[j]
            sims[c_j] += DistMatrix[i][j]
        for c_ in range(l):
            if c_ == c:
                a_i = sims[c_] / (len(clusterResults[c_]) - 1)
            else:
                b_tmp = sims[c_] / (len(clusterResults[c_]))
                if (b_tmp < b_i) or (b_i < 0):
                    b_i = b_tmp

        if(c != 0) and opt1:
            b_i *= optval

        if a_i < b_i:
            s_i = 1 - (a_i / b_i)
        elif a_i == b_i:
            s_i = 0
        else:
            s_i = (b_i / a_i) - 1

        if opt2 and (c != 0) and (s_i < 0):
            s_i = - s_i

        res[c].append(s_i)

    for resitem in res:
        resitem.sort(reverse=True)

    return res


def GetSilhouetteCoeff(Silhouettes):
    res = -2
    for svals in Silhouettes:
        mean = 0
        for s in svals:
            mean += s
        mean = mean / len(svals)
        if (res < -1) or (res > mean):
            res = mean
    return res


def GetSilhouetteMin(Silhouettes):
    res = -2
    for svals in Silhouettes:
        for s in svals:
            if (res < -1) or (res > s):
                res = s
    return res


ExpRootDir = "/home/cadit/WTK/FelineExp/ExpReports2/BlendedMetric"
report = GetReport(ExpRootDir)
hasNoise = report['IsNoiseClusterDefined']
distmat = report['DistanceMatrix']
ExpName = report['ExperimentName']
ClusterNames = report['ClusterNames']
Path_SilhouetteGraphAll = ExpRootDir + \
    "/Report/Images/" + ExpName + "_SilhouettesAll.jpg"
clusterResults = report['ClusteringResults']

Silhouettes = Silhouette_cluster(distmat, clusterResults)

SilhouetteCoeff = GetSilhouetteCoeff(Silhouettes[1:])
SilhouetteMin = GetSilhouetteMin(Silhouettes[1:])

print(SilhouetteCoeff)
print(SilhouetteMin)

EvaluationImage.CreateImage_SilhouetteGraph_All(
    ClusterNames, Silhouettes, Path_SilhouetteGraphAll)
