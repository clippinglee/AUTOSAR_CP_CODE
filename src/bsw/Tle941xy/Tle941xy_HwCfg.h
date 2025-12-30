/*****************************************************************************************************************
******************************************************************************************************************
*  Copyright (C) .                                                                            
*  All rights reserved.                                                                                           
******************************************************************************************************************
*  FileName: Tle941xy                                                                                             
*  Content:  Tle941xy family drive
*  Category: Tle94103 Tle94104 Tle94106 Tle94108 Tle94110 Tle94112 
******************************************************************************************************************
*  Revision Management                                                                                            
*  yyyy.mm.dd    name              version      description                                                       
*  ----------    --------          -------      -----------------------------------                               
*  2022.03.30    clipping            v0001        Frist edit                                                        
******************************************************************************************************************
******************************************************************************************************************/
/* Include Headerfiles  */
#ifndef _TLE941XY_HWCFG_H_
#define _TLE941XY_HWCFG_H_

#include "Tle941xy_Types.h"
#include "Spi.h"
#include "Dio.h"

#define TLE941XY_TLE94103_CHIP_EN STD_OFF
#define TLE941XY_TLE94104_CHIP_EN STD_OFF
#define TLE941XY_TLE94106_CHIP_EN STD_OFF
#define TLE941XY_TLE94108_CHIP_EN STD_OFF
#define TLE941XY_TLE94110_CHIP_EN STD_OFF
#define TLE941XY_TLE94112_CHIP_EN STD_ON


typedef enum
{
    TLE941XY_CHIP_0 = 0u,
    TLE941XY_CHIP_MAX
}Tle941xy_ChipId_e;

typedef enum
{
    TLE941XY_GROUP_0 = 0u,
    TLE941XY_GROUP_1 = 1u,
    TLE941XY_GROUP_MAX
}Tle941xy_GroupId_e;

extern const Tle941xy_ChipType cTle941xy_atChipCfg[TLE941XY_GROUP_MAX][TLE941XY_CHIP_MAX];
extern const Tle941xy_GroupType cTle941xy_atGroupCfg[TLE941XY_GROUP_MAX];
extern const uint8 cTle941xy_au8ChnModeCfg[TLE941XY_GROUP_MAX][TLE941XY_CHIP_MAX][TLE941XY_CHANNEL_MAX];
extern const Tle941xy_PwmType cTle941xy_atChipFmPwmFreqCfg[TLE941XY_GROUP_MAX][TLE941XY_CHIP_MAX];
extern const boolean cTle941xy_abChipFreeWheelingCfg[TLE941XY_GROUP_MAX][TLE941XY_CHIP_MAX][TLE941XY_CHANNEL_MAX];
extern const boolean cTle941xy_abChipHS1And2LedModeCfg[TLE941XY_GROUP_MAX][TLE941XY_CHIP_MAX][2];

#endif