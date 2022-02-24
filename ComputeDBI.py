import json
import os
import sys
from datetime import datetime


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


def DaviesBouldin(DistMatrix, Labels, opt=1):

    l = len(Labels)
    median = GetMedianIndices(DistMatrix, Labels)

    # compute S_i
    S = [0] * l
    for i in range(0, l):
        T = len(Labels[i])
        for j in Labels[i]:
            S[i] = S[i] + DistMatrix[median[i]][j]
        S[i] = S[i] / T * opt

    # compute M
    M = [[0] * l for i in range(0, l)]
    for i in range(0, l):
        for j in range(i+1, l):
            if DistMatrix[median[i]][median[j]] <= 0:
                M[i][j] = M[j][i] = 0.001
            else:
                M[i][j] = M[j][i] = DistMatrix[median[i]][median[j]]

    # compute D_i
    DBI = 0
    for i in range(0, l):
        D_i = 0
        for j in range(0, l):
            if i == j:
                continue
            R_ij = (S[i] + S[j]) / M[i][j]
            if R_ij > D_i:
                D_i = R_ij
        DBI = DBI + D_i
    DBI = DBI / l
    return DBI


def run(ExpRootDir):
    report = GetReport(ExpRootDir)
    hasNoise = report['IsNoiseClusterDefined']
    distmat = report['DistanceMatrix']
    if hasNoise:
        labels = report['ClusteringResults'][1:]
    else:
        labels = report['ClusteringResults']
    return DaviesBouldin(DistMatrix=distmat, Labels=labels)


if __name__ == "__main__":
    out = {}
    for i, arg in enumerate(sys.argv[1:-2]):
        ExpRootDir = arg
        dbi = run(ExpRootDir)
        if not os.path.isdir(ExpRootDir):
            Log("Invalid Feline result root directory: " + ExpRootDir)
            exit()
        else:
            ExpName = ExpRootDir.split("/")[-1]
            out[ExpName] = dbi

    print(out)
