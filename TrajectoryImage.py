import json
import pandas
import numpy as np
import plotly.express as px

MapboxToken = "pk.eyJ1IjoiaGFzaGNvbG9uIiwiYSI6ImNrcXQ5MmphbzA1eHUyb21kOXR0bmFmejcifQ.xamoLzfJsLEryttWEJ0D8g"
imageSize = {
    "width": 1920,
    "height": 1080,
    "zoom": 8
}


def Convert_FeatureCollection2Df(featureCollection, tag):
    # list to collect dataframe of child features
    dflist = []
    # flag to cut down line segments
    linecnt = 0

    # loop child features
    for f in featureCollection['features']:
        # if child is a featurecollection
        if f['type'] == "FeatureCollection":
            df_tmp = Convert_FeatureCollection2Df(f, tag)
        # if child is a feature
        elif f['type'] == "Feature":
            df_tmp = Convert_Feature2Df(f, tag)

        # add current line count to childs
        df_tmp['lineno'] = df_tmp['lineno'] + linecnt
        # renew line count
        linecnt = df_tmp['lineno'].max() + 1
        # collect the results
        dflist.append(df_tmp)
    # merge collected results to a datafrmae
    df = pandas.concat(dflist, ignore_index=True)
    return df


def Convert_Feature2Df(feature, tag):
    lats = []
    lons = []
    tags = []
    line = []
    # works only form linestring type
    if feature['geometry']['type'] == "LineString":
        for pos in feature['geometry']['coordinates']:
            lons.append(pos[0])
            lats.append(pos[1])
            tags.append(tag)
            line.append(0)
    d = {'lon': lons, 'lat': lats, 'tag': tags, 'lineno': line}
    df = pandas.DataFrame(data=d)
    return df


def Convert_Geojson2Df(geoJsonData, tag):
    if geoJsonData['type'] == "FeatureCollection":
        df = Convert_FeatureCollection2Df(geoJsonData, tag)
    elif geoJsonData['type'] == "Feature":
        df = Convert_Feature2Df(geoJsonData, tag)
    return df


def CreateImage_GeoJson(geoJson, filepath):
    # convert geoJson to pandas dataframe
    geoJsonDf = Convert_Geojson2Df(geoJson, "")
    Fig = px.line_mapbox(geoJsonDf, lat="lat", lon="lon",
                         color="tag", line_group='lineno', zoom=imageSize['zoom'])
    Fig.update_layout(mapbox_style="dark", mapbox_accesstoken=MapboxToken)
    Fig.update_layout(margin={"r": 0, "t": 0, "l": 0, "b": 0})
    Fig.write_image(
        filepath, width=imageSize['width'], height=imageSize['height'])
    return Fig


def CreateImage_ManyGeoJson(geoJsonList, tagList, filepath, showtags=True):
    geoJsonDfList = []
    linecnt = 0
    # for all given geoJsons, convert each geojson to dataframe
    for geoJson, tag in zip(geoJsonList, tagList):
        geoJsonDf = Convert_Geojson2Df(geoJson, tag)
        # lineno should be changed before merging
        geoJsonDf['lineno'] = geoJsonDf['lineno'] + linecnt + 1
        # renew line count
        linecnt = geoJsonDf['lineno'].max()
        geoJsonDfList.append(geoJsonDf)
    df = pandas.concat(geoJsonDfList, ignore_index=True)
    Fig = px.line_mapbox(df, lat="lat", lon="lon",
                         color="tag", line_group='lineno', zoom=imageSize['zoom'])
    Fig.update_layout(mapbox_style="dark", mapbox_accesstoken=MapboxToken)
    Fig.update_layout(margin={"r": 0, "t": 0, "l": 0, "b": 0})
    if showtags == False:
        Fig.update_traces(showlegend=False)
    Fig.write_image(
        filepath, width=imageSize['width'], height=imageSize['height'])
    return Fig


def CreateImage_GeoJsonFilePath(geoJsonFilePath, imgFilepath):
    geoJson = open(geoJsonFilePath, 'r')
    geoJson = json.load(geoJson)
    return CreateImage_GeoJson(geoJson, imgFilepath)


def CreateImage_ManyGeoJsonFilePath(geoJsonFilePathList, tagList, imgFilepath, showtags=True):
    geoJsonList = []
    for geoJsonFilePath in geoJsonFilePathList:
        geoJson = open(geoJsonFilePath, 'r')
        geoJson = json.load(geoJson)
        geoJsonList.append(geoJson)
    return CreateImage_ManyGeoJson(geoJsonList, tagList, imgFilepath, showtags)


def CreateImage_RouteLists(routeList, tagList, imgFilepath, showtags=True):
    geoJsonDfList = []
    linecnt = 0
    for route, tag in zip(routeList, tagList):
        tmpGeoJsonFilePath = route[0]
        routeIdx = route[1]
        tmpGeoJson = open(tmpGeoJsonFilePath, 'r')
        tmpGeoJson = json.load(tmpGeoJson)

        tmpGeoJsonDf = Convert_Geojson2Df(tmpGeoJson, tag)
        # fileter geojson by route index
        tmpGeoJsonDf = tmpGeoJsonDf[tmpGeoJsonDf['lineno'] == routeIdx]
        # change lineno
        tmpGeoJsonDf['lineno'] = linecnt
        linecnt += 1
        # append to geoJsonDflist
        geoJsonDfList.append(tmpGeoJsonDf)

    df = pandas.concat(geoJsonDfList, ignore_index=True)
    Fig = px.line_mapbox(df, lat="lat", lon="lon",
                         color="tag", line_group='lineno', zoom=imageSize['zoom'])
    Fig.update_layout(mapbox_style="dark", mapbox_accesstoken=MapboxToken)
    Fig.update_layout(margin={"r": 0, "t": 0, "l": 0, "b": 0})
    if showtags == False:
        Fig.update_traces(showlegend=False)
    Fig.write_image(
        imgFilepath, width=imageSize['width'], height=imageSize['height'])
    return Fig
