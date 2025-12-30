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
#include "Tle941xy_Types.h"
#include "Tle941xy_HwCfg.h"

uint8 gTle941xy_u8Group0ChipNum = TLE941XY_CHIP_MAX;
const Tle941xy_GroupType cTle941xy_atGroupCfg[TLE941XY_GROUP_MAX] =
{
    {
        SpiConf_SpiChannel_SpiChannel_TLE94112_0, 
        SpiConf_SpiSequence_SpiSequence_TLE94112_0,
        TLE941XY_DAISY_CHAIN_NO_USER,
        &gTle941xy_u8Group0ChipNum,
    },
    {
        SpiConf_SpiChannel_SpiChannel_TLE94112_1, 
        SpiConf_SpiSequence_SpiSequence_TLE94112_1,
        TLE941XY_DAISY_CHAIN_NO_USER,
        &gTle941xy_u8Group0ChipNum,
    },
};

const Tle941xy_ChipType cTle941xy_atChipCfg[TLE941XY_GROUP_MAX][TLE941XY_CHIP_MAX] = 
{
    {
        {
            TLE941XY_TLE94112,
            DioConf_DioChannel_DioChannel_P24_00,
        },
    },
    {
        {
            TLE941XY_TLE94112,
            DioConf_DioChannel_DioChannel_P24_01,
        },
    },
};


const uint8 cTle941xy_au8ChnModeCfg[TLE941XY_GROUP_MAX][TLE941XY_CHIP_MAX][TLE941XY_CHANNEL_MAX] =
{
    {
        {
            /* TLE94112*/
            TLE941XY_CHN_CTRL_DO,
            TLE941XY_CHN_CTRL_DO,
            TLE941XY_CHN_CTRL_DO,
            TLE941XY_CHN_CTRL_DO,

            TLE941XY_CHN_CTRL_DO,
            TLE941XY_CHN_CTRL_DO,
            TLE941XY_CHN_CTRL_DO,
            TLE941XY_CHN_CTRL_DO,

            TLE941XY_CHN_CTRL_DO,
            TLE941XY_CHN_CTRL_DO,
            TLE941XY_CHN_CTRL_DO,
            TLE941XY_CHN_CTRL_DO,
        },
    },
    {
        {
            /* TLE94112*/
            TLE941XY_CHN_CTRL_DO,
            TLE941XY_CHN_CTRL_DO,
            TLE941XY_CHN_CTRL_DO,
            TLE941XY_CHN_CTRL_DO,

            TLE941XY_CHN_CTRL_DO,
            TLE941XY_CHN_CTRL_DO,
            TLE941XY_CHN_CTRL_DO,
            TLE941XY_CHN_CTRL_DO,

            TLE941XY_CHN_CTRL_DO,
            TLE941XY_CHN_CTRL_DO,
            TLE941XY_CHN_CTRL_DO,
            TLE941XY_CHN_CTRL_DO,
        },
    },
};

const Tle941xy_PwmType cTle941xy_atChipFmPwmFreqCfg[TLE941XY_GROUP_MAX][TLE941XY_CHIP_MAX] =
{
    {
        {
            TLE941XY_FM_MOD_FREQ_62KHZ,
            TLE941XY_PWM_OFF,
            TLE941XY_PWM_OFF,
            TLE941XY_PWM_OFF,
        },
    },
    {
        {
            TLE941XY_FM_MOD_FREQ_62KHZ,
            TLE941XY_PWM_OFF,
            TLE941XY_PWM_OFF,
            TLE941XY_PWM_OFF,
        },
    }
};

const boolean cTle941xy_abChipFreeWheelingCfg[TLE941XY_GROUP_MAX][TLE941XY_CHIP_MAX][TLE941XY_CHANNEL_MAX] =
{
    {
        {
            /* TLE94112*/
            TLE941XY_CHN_FW_OFF,
            TLE941XY_CHN_FW_OFF,
            TLE941XY_CHN_FW_OFF,
            TLE941XY_CHN_FW_OFF,

            TLE941XY_CHN_FW_OFF,
            TLE941XY_CHN_FW_OFF,
            TLE941XY_CHN_FW_OFF,
            TLE941XY_CHN_FW_OFF,

            TLE941XY_CHN_FW_OFF,
            TLE941XY_CHN_FW_OFF,
            TLE941XY_CHN_FW_OFF,
            TLE941XY_CHN_FW_OFF,
        },
    },
    {
        {
            /* TLE94112*/
            TLE941XY_CHN_FW_OFF,
            TLE941XY_CHN_FW_OFF,
            TLE941XY_CHN_FW_OFF,
            TLE941XY_CHN_FW_OFF,

            TLE941XY_CHN_FW_OFF,
            TLE941XY_CHN_FW_OFF,
            TLE941XY_CHN_FW_OFF,
            TLE941XY_CHN_FW_OFF,

            TLE941XY_CHN_FW_OFF,
            TLE941XY_CHN_FW_OFF,
            TLE941XY_CHN_FW_OFF,
            TLE941XY_CHN_FW_OFF,
        },
    },
};

const boolean cTle941xy_abChipHS1And2LedModeCfg[TLE941XY_GROUP_MAX][TLE941XY_CHIP_MAX][2] =
{
    {
        {
            TLE941XY_CHN_LED_OFF,
            TLE941XY_CHN_LED_OFF,
        },
    },
    {
        {
            TLE941XY_CHN_LED_OFF,
            TLE941XY_CHN_LED_OFF,
        },
    },
};