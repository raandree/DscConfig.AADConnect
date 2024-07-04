//********************************************************
//*                                                      *
//*   Copyright (C) Microsoft. All rights reserved.      *
//*                                                      *
//********************************************************
//////////////////////////////////////////////////////////////////////////////
//
//      Performance counters for Microsoft Azure AD Sync 
//      Synchronization Service.
//
//////////////////////////////////////////////////////////////////////////////

#pragma once

#define PRFOBJ_MMS_CS                                        0
#define MMSPERF_CS_FIX_BACKLINKS_OBJECTS                     2
#define MMSPERF_CS_FIX_BACKLINKS_OBJECTS_RATE                4
#define MMSPERF_CS_FIX_BACKLINKS_LINKS                       6
#define MMSPERF_CS_FIX_BACKLINKS_LINKS_RATE                  8
#define MMSPERF_CS_PRUNE_CACHE_READS                        10
#define MMSPERF_CS_PRUNE_CACHE_READS_RATE                   12
#define MMSPERF_CS_PRUNE_CACHE_WRITES                       14
#define MMSPERF_CS_PRUNE_CACHE_WRITES_RATE                  16
#define MMSPERF_CS_OBJECTS_PRUNED                           18
#define MMSPERF_CS_OBJECTS_PRUNED_RATE                      20
#define MMSPERF_CS_PRUNE_RECURSION_LEVEL                    22

#define PRFOBJ_MMS_MA                                       24
#define MMSPERF_MA_OBJECTS_IMPORTED                         26
#define MMSPERF_MA_OBJECTS_IMPORTED_RATE                    28
#define MMSPERF_MA_OBJECTS_EXPORTED                         30
#define MMSPERF_MA_OBJECTS_EXPORTED_RATE                    32
#define MMSPERF_MA_EXTENSION_WRITE_EXPORT_FILE_TIMER        34
#define MMSPERF_MA_EXTENSION_DELIVER_EXPORT_FILE_TIMER      36
#define MMSPERF_MA_EXTENSION_CALL_BASED_EXPORT_TIMER        38
#define MMSPERF_MA_EXTENSION_GENERATE_IMPORT_FILE_TIMER     40

#define PRFOBJ_MMS_SE                                       42
#define MMSPERF_SE_RETRYS_PROCESSED                         44
#define MMSPERF_SE_RETRYS_PROCESSED_RATE                    46

#define PRFOBJ_MMS_HS                                       48
#define MMSPERF_HS_REBUILD_LINK_TIMER                       50 // Time in Rebuild_Links in CS Object
#define MMSPERF_HS_CS_PERSIST_TIMER                         52 // Time in Persist in CS Object
#define MMSPERF_HS_UPDATE_ANCHOR_TIMER                      54
#define MMSPERF_HS_PUT_ANCHOR_TIMER                         56
#define MMSPERF_HS_OBJECTS_CONVERTED_TIMER                  58 // Time in MA to convert object to image
#define MMSPERF_HS_STAGE_TIMER                              60 // Time in stage an object including MMSPERF_HS_STAGE_CREATE_TIMER,MMSPERF_HS_STAGE_ADD_DIMAGE_TIMER,MMSPERF_HS_SETUP_LINKS_TIMER,MMSPERF_HS_CS_PERSIST_TIMER
#define MMSPERF_HS_CS_TO_MV_TIMER                           62 // Time in IAF including MMSPERF_HS_SETUP_LINKS_TIMER
#define MMSPERF_HS_SETUP_LINKS_TIMER                        64 // Time in SetupLinks in CS
#define MMSPERF_HS_PROVISION_TIMER                          66 // Time in Provisioning
#define MMSPERF_HS_MV_TO_CS_TIMER                           68 // Time in EAF including MMSPERF_HS_CS_PERSIST_TIMER
#define MMSPERF_HS_MV_PERSIST_TIMER                         70 // Time to persist MV object
#define MMSPERF_HS_STAGE_CREATE_TIMER                       72 // Time to create a stanging image
#define MMSPERF_HS_RETRY_TIMER                              74 // Time in retry including ...
#define MMSPERF_HS_STAGE_ADD_DIMAGE_TIMER                   76 // Time to add a DImage to pending import
#define MMSPERF_HS_FIND_CS_OBJECT_TIMER                     78 // Time to find a CS Object for Import + Sync
#define MMSPERF_HS_OBJECTS_READ_TIMER                       80 // Time in MA to read object, may overlap with MMSPERF_HS_OBJECTS_CONVERTED_TIMER
#define MMSPERF_HS_SYNCHRONIZE_TIMER                        82 // Time in sync, including MMSPERF_HS_FIND_CS_OBJECT_TIMER, MMSPERF_HS_LINK_MV_TIMER, MMSPERF_HS_CS_TO_MV_TIMER, MMSPERF_HS_PROVISION_TIMER , MMSPERF_HS_MV_TO_CS_TIMER, MMSPERF_HS_MV_PERSIST_TIMER
#define MMSPERF_HS_LINK_MV_TIMER                            84 // Time in linking to MV including join + project
#define MMSPERF_HS_GET_EXPORT_IMAGE_TIMER                   86 // Tiem in Loading Image for export
#define MMSPERF_HS_EXPORT_PARENTS_TIMER                     88 // Time in exporting parents
#define MMSPERF_HS_EXPORT_TO_CD_TIMER                       90 // Time in export to connected directory
#define MMSPERF_HS_POST_EXPORT_PROC_TIMER                   92 // Time in processing export updates (anchors, failures, etc...)
#define MMSPERF_HS_PUT_IMPORTS_TIMER                        94 // PutImports (syncimport.cpp)
#define MMSPERF_HS_IMPORT_START_TRANS_TIMER                 96 // Import->Start Transaction
#define MMSPERF_HS_IMPORT_END_TRANS_TIMER                   98 // Import->End Transaction
#define MMSPERF_HS_SYNC_START_TRANS_TIMER                  100 // Sync->Start Transaction
#define MMSPERF_HS_SYNC_END_TRANS_TIMER                    102 // Sync->End Transaction
#define MMSPERF_HS_OBJECTS_PRUNED_TIMER                    104 // Object Whacking
#define MMSPERF_HS_PRE_EXPORT_PROC_TIMER                   106 // Time in processing pre-export updates (anchors, failures, etc...)
#define MMSPERF_HS_STAMP_ANCHOR_TIMER                      108 // Time to stamp anchor on CS/Tower
#define MMSPERF_HS_ESCROW_CHANGE_TIMER                     110 // Time to push pending to escrowed

#define PERF_MAX_COUNTER                                   110 // Note: same as highest counter value

