/*****************************************************************************************************************
******************************************************************************************************************
*  Copyright (C) .                                                                            
*  All rights reserved.                                                                                           
******************************************************************************************************************
*  FileName: Tle9210x                                                                                             
*  Content:  Tle9210x family drive
*  Category: Tle92104 Tle92108
******************************************************************************************************************
*  Revision Management                                                                                            
*  yyyy.mm.dd    name              version      description                                                       
*  ----------    --------          -------      -----------------------------------                               
*  2022.03.30    clipping            v0001        Frist edit                                                        
******************************************************************************************************************
******************************************************************************************************************/
/* Include Headerfiles  */
#ifndef _TLE9210X_HWCFG_H_
#define _TLE9210X_HWCFG_H_

#include "Tle9210x_Types.h"

typedef enum
{
    TLE9210X_CHIP_0 = 0u,
    TLE9210X_CHIP_MAX
}Tle9210x_ChipId_e;

typedef enum
{
    TLE9210X_GROUP_0 = 0u,
    TLE9210X_GROUP_1 = 1u,
    TLE9210X_GROUP_2 = 2u,
    TLE9210X_GROUP_MAX
}Tle9210x_GroupId_e;

#define TLE9210X_TLE92104_CHIP_EN STD_OFF
#define TLE9210X_TLE92108_CHIP_EN STD_ON


extern const Tle9210x_GroupType cTle9210x_atGroupCfg[TLE9210X_GROUP_MAX];
extern const Tle9210x_ChipType cTle9210x_atChipCfg[TLE9210X_GROUP_MAX][TLE9210X_CHIP_MAX];
extern const Tle9210x_PwmChnType cTle9210x_atPwmChnCfg[TLE9210X_GROUP_MAX][TLE9210X_CHIP_MAX][TLE9210X_PWM_CHN_MAX];
extern const Tle9210x_HbChnType cTle9210x_atHbChnCfg[TLE9210X_GROUP_MAX][TLE9210X_CHIP_MAX][TLE9210X_HB_CHN_MAX];

#endif
