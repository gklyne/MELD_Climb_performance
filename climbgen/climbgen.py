"""
Generate Climb! MELD data from spreatsheet
"""

__author__      = "Graham Klyne (GK@ACM.ORG)"
__copyright__   = "Copyright 2017, G. Klyne"
__license__     = "MIT (http://opensource.org/licenses/MIT)"

import sys
import os
import os.path
# import urlparse
# import shutil
import json
# import errno

import logging
log = logging.getLogger(__name__)

from grid.grid import GridExcel

def open_spreadsheet(name):
    g = GridExcel(name)
    return g

def open_json(dirname, filename):
    with open(dirname+"/"+filename) as inpstr:
        jsondata = json.load(inpstr)
    return jsondata

def write_json(dirname, filename, jsondata):
    log.info("write_json: %s/%s"%(dirname, filename))    
    try:
        os.makedirs(dirname)
    except OSError:
        pass
    with open(dirname+"/"+filename, "w") as outstr:
        json.dump(jsondata, outstr, sort_keys=True, indent=2, separators=(',', ': '))
    return

def col_index(name):
    """
    Map column name (per Excel) to index
    """
    col_names = (
        [ "A", "B", "C", "D", "E", "F", "G", "H", "I", "J"  # stage
        , "K", "L", "M", "N", "O", "P", "Q", "R", "S"       # auto
        , "T", "U", "V", "W", "X", "Y", "Z", "AA", "AB"     # mc1
        , "AC", "AD", "AE", "AF", "AG", "AH", "AI"          # mc2
        , "AJ", "AK", "AL", "AM", "AN", "AO", "AP"          # mc3
        , "AQ", "AR", "AS", "AT", "AU", "AV", "AW"          # mc4
        , "AX", "AY", "AZ", "BC", "BD"                      # mc5
        , "BE", "BF", "BG", "BH", "BI"                      # unused?
        ])
    return col_names.index(name)

def get_col_index(hdr, name, start=0, end=9999):
    """
    Find index of named column in range [start..end], or -1
    """
    j = start
    try:
        while j < end:
            # log.debug("get_col_index try %d, %s == %s"%(j, hdr[j], name))
            if hdr[j] == name:
                return j
            j += 1
    except IndexError, e:
        pass
    return -1

def get_col(hdr, row, name, start=0, end=9999):
    j = get_col_index(hdr, name, start=start, end=end)
    if j >= 0:
        return row[j]
    return None

def analyze_table_data(table):
    data = { "stages": [] }
    hdr  = table[0]
    # log.debug("hdr %r"%(hdr,))
    for row in table.rows(1):
        # log.debug("row %r"%(row,))
        # For each stage-related field...
        stage_id  = get_col(hdr, row, "stage")
        if not stage_id:
            break
        # log.info("stage_id: %r"%(stage_id,))
        stagedata = (
            { "stage":          stage_id
            , "next":           get_col(hdr, row, "next")
            , "meifile":        get_col(hdr, row, "meifile")
            , "no_effect":      get_col(hdr, row, "no_effect")
            , "rain_effect":    get_col(hdr, row, "rain_effect")
            , "snow_effect":    get_col(hdr, row, "snow_effect")
            , "wind_effect":    get_col(hdr, row, "wind_effect")
            , "storm_effect":   get_col(hdr, row, "storm_effect")
            , "sun_effect":     get_col(hdr, row, "sun_effect")
            , "default_cue":    get_col(hdr, row, "default_cue")
            , "auto_actions":   (
                { "name":           None
                , "cue":            None
                , "midi":           None
                , "midi2":          None
                , "delay":          None
                , "monitor":        None
                , "v.animate":      None
                , "v.background":   None
                , "v.mc":           None
                , "v.mc.delay":     None
                , "app":            None
                })
            , "mc_actions":     [None]
            })
        # log.debug("stage %(stage)s"%stagedata)
        # Auto actions
        auto_beg = col_index("K")
        auto_end = col_index("S")
        auto_actions = (
            { "mc":           None
            , "mc_hdr":       None
            , "name":         None
            , "cue":          get_col(hdr, row, "cue",          start=auto_beg, end=auto_end)
            , "midi":         get_col(hdr, row, "midi",         start=auto_beg, end=auto_end)
            , "midi2":        get_col(hdr, row, "midi2",        start=auto_beg, end=auto_end)
            , "delay":        get_col(hdr, row, "delay",        start=auto_beg, end=auto_end)
            , "monitor":      get_col(hdr, row, "monitor",      start=auto_beg, end=auto_end)
            , "v.animate":    get_col(hdr, row, "v.animate",    start=auto_beg, end=auto_end)
            , "v.background": get_col(hdr, row, "v.background", start=auto_beg, end=auto_end)
            , "v.mc":         get_col(hdr, row, "v.mc",         start=auto_beg, end=auto_end)
            , "v.mc.delay":   get_col(hdr, row, "v.mc.delay",   start=auto_beg, end=auto_end)
            , "app":          get_col(hdr, row, "app",          start=auto_beg, end=auto_end)
            })
        stagedata["auto_actions"] = auto_actions
        # Muzicode-triggered actions
        muzicode_groups = (
            { 1:  get_col_index(hdr, "mc1:")
            , 2:  get_col_index(hdr, "mc2:")
            , 3:  get_col_index(hdr, "mc3:")
            , 4:  get_col_index(hdr, "mc4:")
            , 5:  get_col_index(hdr, "mc5:")
            })
        for mcg in muzicode_groups:
            mc_beg = muzicode_groups[mcg]
            mc_end = muzicode_groups.get(mcg+1, 9999)
            mc_actions = (
                { "mc_hdr":       hdr[mc_beg]
                , "mc":           row[mc_beg]
                , "name":         get_col(hdr, row, "name",       start=mc_beg, end=mc_end)
                , "cue":          get_col(hdr, row, "cue",        start=mc_beg, end=mc_end)
                , "midi":         get_col(hdr, row, "midi",       start=mc_beg, end=mc_end)
                , "monitor":      get_col(hdr, row, "monitor",    start=mc_beg, end=mc_end)
                , "midi2":        None
                , "delay":        None
                , "v.animate":    get_col(hdr, row, "v.animate",  start=mc_beg, end=mc_end)
                , "v.background": None
                , "v.mc":         get_col(hdr, row, "v.mc",       start=mc_beg, end=mc_end)
                , "v.mc.delay":   get_col(hdr, row, "v.mc.delay", start=mc_beg, end=mc_end)
                , "app":          get_col(hdr, row, "app",        start=mc_beg, end=mc_end)
                })
            stagedata["mc_actions"].append(mc_actions)
        data["stages"].append(stagedata)
    return data

def make_id(type_id, entity_id):
    if entity_id:
        return "%s/%s"%(type_id, entity_id)
    return ""

def generate_meld_data(data, jsondata, base_dir):
    # See:
    #   https://github.com/oerc-music/meld/blob/master/server/generate_climb_scores.py 
    #   https://github.com/cgreenhalgh/fast-performance-demo/tree/master/scoretools
    #
    status = 0
    stage_json = {}
    for sj in jsondata:
        stage_json[sj["stage"]]= sj 
    for stage in data["stages"]:
        log.info("generate_meld_data: stage %(stage)s"%stage)
        # Generate climb_Stage_Score entity
        stage_id              = stage["stage"]
        stage_ref             = make_id("climb_Stage_Score", stage_id)
        published_score_id    = stage_id+"_published"
        published_score_ref   = make_id("climb_Stage_Published_Score", published_score_id)
        published_score_label = "Published score for \"%s\""%(stage_id,)
        auto_actions_id       = stage_id+"_auto"
        auto_actions_ref      = make_id("climb_Actions", auto_actions_id)
        stage_score = (
            { "@context": [
                {
                  "@base": "../../"
                },
                "../../coll_context.jsonld"
              ],
              "@id": make_id("climb_Stage_Score", stage_id),
              "@type": [
                "climb:Stage_Score",
                "frbr:Group_1_entity",
                "mo:MusicalExpression",
                "mo:Score",
                "frbr:Expression",
                "annal:EntityData"
              ],
              "annal:id":                   stage_id,
              "annal:type":                 "climb:Stage_Score",
              "annal:type_id":              "climb_Stage_Score",
              "rdfs:comment":               
                    "# Climb! %s\r\n\r\nStage %s of a Climb! performance."%(stage_id, stage_id),
              "rdfs:label":                 "Climb! %s"%(stage_id,),
              "mo:published_as":            published_score_ref,
              "climb:next_stage":           make_id("climb_Stage_Score", stage["next"]),
              "climb:default_cue_stage":    make_id("climb_Stage_Score", stage["default_cue"]),
              "climb:no_effect":            stage["no_effect"],
              "climb:rain_effect":          stage["rain_effect"],
              "climb:snow_effect":          stage["snow_effect"],
              "climb:wind_effect":          stage["wind_effect"],
              "climb:storm_effect":         stage["storm_effect"],
              "climb:sun_effect":           stage["sun_effect"],
              "climb:auto":                 auto_actions_id,
              "frbr:part":                  []
            })
        # Generate published score record
        stage_published_score = (
            { "@context": [
                {
                  "@base": "../../"
                },
                "../../coll_context.jsonld"
              ],
              "@id":                        published_score_ref,
              "@type": [
                "meld:Published_Score_MEI",
                "mo:MusicalManifestation",
                "frbr:Manifestation_MEI",
                "frbr:Manifestation",
                "meld:Manifestation_MEI",
                "mo:PublishedScore",
                "frbr:Group_1_entity",
                "meld:Manifestation",
                "annal:EntityData"
              ],
              "annal:id":                   published_score_id,
              "annal:type":                 "",
              "annal:type":                 "climb:Stage_Published_Score",
              "annal:type_id":              "climb_Stage_Published_Score",
              "rdfs:label":                 published_score_label,
              "rdfs:comment":               "# %s\r\n\r\n"%(published_score_label,),
              "frbr:url":                   "climbstage:%s"%(stage["meifile"],),
            })
        # Generate auto actions record ("climb_Actions/<stage>_auto")
        generate_actions(
            auto_actions_id, auto_actions_ref, 
            "Stage %s auto actions"%(stage_id,), 
            stage["auto_actions"]
            )
        # Generate Muzicode descriptions; add links as "frbr:part {"@id": ...} values
        mc_json = {}
        for mj in stage_json[stage_id]["mcs"]:
            mc_json[mj["name"]] = mj
        for mc in stage["mc_actions"]:
            mc_ref  = generate_muzicode_data(stage_id, stage["meifile"], mc, mc_json)
            mc_part = { "@id": mc_ref }
            if mc_ref:
                stage_score["frbr:part"].append(mc_part)
        # Write out stage data (locally - ready to copy later)
        outdir = "d/"+stage_ref
        write_json(outdir, "entity_data.jsonld", stage_score)
        # Write out published score description (locally - ready to copy later)
        outdir = "d/"+published_score_ref
        write_json(outdir, "entity_data.jsonld", stage_published_score)
    return status

def generate_actions(actions_id, actions_ref, actions_label, actions_data):
    # See: https://github.com/cgreenhalgh/fast-performance-demo/tree/master/scoretools
    actions_json = (
        {
          "@context": [
            {
              "@base": "../../"
            },
            "../../coll_context.jsonld"
          ],
          "@id":   actions_ref,
          "@type": [
            "climb:Actions",
            "annal:EntityData"
          ],
          "annal:id":                           actions_id,
          "annal:type":                         "climb:Actions",
          "annal:type_id":                      "climb_Actions",
          "rdfs:comment":                       "# " + actions_label,
          "rdfs:label":                         actions_label,
          "climb:cue_stage":                    actions_data["cue"],
          "climb:action_midi":                  actions_data["midi"],
          "climb:action_midi2_delayed":         actions_data["midi2"],
          "climb:action_midi2_delay_value":     actions_data["delay"],
          "climb:action_monitor_visual":        actions_data["monitor"],
          "climb:action_background_visual":     actions_data["v.background"],
          "climb:action_animation":             actions_data["v.animate"],
          "climb:action_animation_delay_value": actions_data["delay"], # 'vdelta' not currently used
          "climb:action_mc_visual":             actions_data["v.mc"],
          "climb:action_mc_delay":              actions_data["v.mc.delay"],
          "climb:action_app_message":           actions_data["app"]
        })
    # Write out actions data (locally - ready to copy later)
    outdir = "d/"+actions_ref
    write_json(outdir, "entity_data.jsonld", actions_json)
    return

def generate_muzicode_data(stage_id, stage_meifile, mc_actions, mc_actions_json):
    # See:
    #   https://github.com/oerc-music/meld/blob/master/server/generate_climb_scores.py 
    #   https://github.com/cgreenhalgh/fast-performance-demo/tree/master/scoretools
    #
    # log.debug("generate_muzicode_data: %s: %r"%(stage_id, mc_actions))
    mc_ref  = None
    mc_name = None
    if mc_actions:
        mc_name = mc_actions["name"]
        if not mc_name and mc_actions["midi"]:
            mc_name = "%s_%s"%(stage_id, mc_actions["mc_hdr"])
    if mc_name:
        mc_id          = mc_name
        mc_ref         = make_id("climb_Muzicode", mc_id)
        mc_label       = "%s: Muzicode %s"%(stage_id, mc_name)
        mc_cue         = make_id("climb_Stage_Score", mc_actions["cue"])
        mc_type        = "CHOICE"
        mc_mei_ref     = None
        mc_mei_label   = mc_label
        mc_meielements = []
        # Extract JSON description of Muzicode if present:
        # this has information not present in the spreadsheet data
        if mc_name in mc_actions_json:
            if "app" in mc_actions_json[mc_name]:
                mc_label   = mc_actions_json[mc_name]["app"]
            else:
                mc_label   = "Muzicode %s (%s)"%(mc_name, mc_type)
            mc_label       = "%s: %s"%(stage_id, mc_label)
            mc_type        = mc_actions_json[mc_name]["type"].upper()
            mc_mei_ref     = make_id("meld_Manifestation_Bag", mc_id)
            mc_mei_label   = mc_label
            mc_meielements = mc_actions_json[mc_name]["meielements"]
        mc_type_ref    = make_id("climb_Muzicode_Type", mc_type)
        # Generate Muzicode description
        mc_json  = (
            {
              "@context": [
                {
                  "@base": "../../"
                },
                "../../coll_context.jsonld"
              ],
              "@id":                    mc_ref,
              "@type": [
                "climb:Muzicode",
                "mo:MusicalExpression",
                "meld:Muzicode",
                "frbr:Group_1_entity",
                "meld:Expression",
                "frbr:Expression",
                "annal:EntityData"
              ],
              "annal:id":               mc_id,
              "annal:type":             "climb:Muzicode",
              "annal:type_id":          "climb_Muzicode",
              "rdfs:label":             mc_label,
              "rdfs:comment":           "# %s\r\n\r\n\r\n\r\n"%(mc_label,),
              "climb:cue_stage":        mc_cue,
              "frbr:embodiment":        mc_mei_ref,
              "mc:type":                mc_type_ref
            })
        # Write out muzicode data (locally - ready to copy later)
        outdir = "d/"+mc_ref
        write_json(outdir, "entity_data.jsonld", mc_json)
        # Generate description of Muzicide embodiment in MEI, if defined
        # (@@ Some "Muzicodes" apear to just generate MIDO outputs with no other aossciated data)
        if mc_mei_ref:
            def meielement(e):
                return { "@id": stage_meifile + e }
            mc_mei_json = (
                {
                  "@context": [
                    {
                      "@base": "../../"
                    },
                    "../../coll_context.jsonld"
                  ],
                  "@id":            mc_mei_ref,
                  "@type": [
                    "meld:Manifestation_Bag",
                    "frbr:Manifestation",
                    "frbr:Group_1_entity",
                    "annal:EntityData"
                  ],
                  "annal:id":       mc_id,
                  "annal:type":     "meld:Manifestation_Bag",
                  "annal:type_id":  "meld_Manifestation_Bag",
                  "rdfs:label":     mc_mei_label,
                  "rdfs:comment":   "# %s\r\n\r\n"%(mc_mei_label,),
                  "rdfs:member":    [ meielement(e) for e in mc_meielements ]
                })
            # Write out muzicode MEI embodiment description (locally - ready to copy later)
            outdir = "d/"+mc_mei_ref
            write_json(outdir, "entity_data.jsonld", mc_mei_json)
        # Generate actions description
        actions_ref   = make_id("climb_Actions", mc_id)
        actions_label = mc_label
        generate_actions(mc_id, actions_ref, actions_label, mc_actions)
        # Generate event template
        # NOTE: MELD does not define a role for event annotations not within a MELD session.
        #       Event annotations not within a MELD session are here arbitrarily considered to be 
        #       templates to be used when generating events within a session.  This is currently a 
        #       modelling artifact, but could in principle be used in actual implementations.
        event_ref     = make_id("climb_Annotation", mc_id)
        event_label   = mc_label
        event_json    = (
            { "@context": [
                {
                  "@base": "../../"
                },
                "../../coll_context.jsonld"
              ],
              "@id":                event_ref,
              "@type": [
                "climb:Annotation",
                "ao:Annotation",
                "meld:Annotation",
                "annal:EntityData"
              ],
              "annal:id":           mc_id,
              "annal:type":         "climb:Annotation",
              "annal:type_id":      "climb_Annotation",
              "rdfs:label":         event_label,
              "rdfs:comment":       "# %s\r\n\r\n\r\n\r\n"%(event_label,),
              "ao:hasBody":         actions_ref,
              "ao:hasTarget":       mc_ref,
              "ao:motivation":      mc_type_ref
            })
        # Write out event data (locally - ready to copy later)
        outdir = "d/"+event_ref
        write_json(outdir, "entity_data.jsonld", event_json)
    return mc_ref

def generate_climb_meld(configbase, argv):
    """
    Top-level logic for MELD generation from spreadsheet data
    """
    climb_table = open_spreadsheet("mkGameEngine2.xlsx")
    climb_data  = analyze_table_data(climb_table)
    climb_json  = open_json(".", "mkGameEngine2.json")
    status = generate_meld_data(climb_data, climb_json, os.path.join(configbase, "d/"))
    return status

def runMain():
    """
    Main program transfer function for setup.py console script
    """
    # configbase = os.path.expanduser("~")
    configbase = os.getcwd()
    return generate_climb_meld(configbase, sys.argv)

if __name__ == "__main__":
    """
    Program invoked from the command line.
    """
    # logging.basicConfig(filename='example.log', filemode='w', level=logging.DEBUG)
    logging.basicConfig(level=logging.DEBUG)
    status = runMain()
    sys.exit(status)

#   /Users/graham/annalist_site/c/MELD_Climb_performance/climbgen
