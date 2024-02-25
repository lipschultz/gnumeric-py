#!/usr/bin/env python

import json
import os
from hatchling.metadata.plugin.interface import MetadataHookInterface

class JSONMetaDataHook(MetadataHookInterface):
    PLUGIN_NAME='JSONMeta'

    def update(self, metadata):
        here = os.path.abspath(os.path.dirname(__file__))
        # it would be nice to make src_file configurable
        src_file = os.path.join(here, "gnumeric", ".constants.json")
        with open(src_file) as src:
            constants = json.load(src)
            metadata["authors"] = [{"name": constants["__author__"], 
                                   "email": constants["__author_email__"] }]
            # next doesn't currently exist
            maintainer_name = constants.get("__maintainer__")
            if not maintainer_name:
                maintainer_name = constants["__author__"]
            metadata["maintainers"] = [{"name": maintainer_name,
                                       "email": constants["__maintainer_email__"]}]
            metadata["license"] = constants["__license__"]
            # since hatch gives version special handling it's not clear what this will do
            metadata["version"] = constants["__version__"]
            # skipped since in separate section of pyrpoject.toml:
            #  constants["__url__"]
 