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
#include "Tle9210x.h"
#include "PFM.h"
#include "Spi.h"
#include "LiBool.h"
#include "Pwm.h"

static boolean sTle9210x_abREGBANKSts[TLE9210X_GROUP_MAX][TLE9210X_CHIP_MAX];
static uint8 sTle9210x_au8GlobalStatus[TLE9210X_GROUP_MAX][TLE9210X_CHIP_MAX];
static uint8 sTle9210x_au8PwmDuty[TLE9210X_GROUP_MAX][TLE9210X_CHIP_MAX][TLE9210X_PWM_CHN_MAX];
static uint8 sTle9210x_au8HbOutSts[TLE9210X_GROUP_MAX][TLE9210X_CHIP_MAX][TLE9210X_HB_CHN_MAX];
static PFM_DefectReportState_t sTle9210x_atDiagResult[TLE9210X_GROUP_MAX][TLE9210X_CHIP_MAX][TLE9210X_HB_CHN_MAX];

static uint8 sTle9210x_au8ChipMode[TLE9210X_GROUP_MAX][TLE9210X_CHIP_MAX];
static uint16 sTle9210x_au16GenCtrl1[TLE9210X_GROUP_MAX][TLE9210X_CHIP_MAX];
static uint16 sTle9210x_au16GenCtrl2[TLE9210X_GROUP_MAX][TLE9210X_CHIP_MAX];

static Tle9210x_GenStsRegType sTle9210x_atGenStsReport[TLE9210X_GROUP_MAX][TLE9210X_CHIP_MAX];

static void Tle9210x_WriteReg(uint8 u8GroupId,uint8* pu8RegBuf, uint16* pu16WtData);
static void Tle9210x_ReadReg( uint8 u8GroupId,uint8* pu8RegBuf,uint16* pu16ReadBuf);
static void Tle9210x_SetChipMode(uint8 u8GroupId,uint8 u8Mode);
static void Tle9210x_GetChipMode(uint8 u8GroupId,uint8 u8ChipId,uint8* pu8Mode);
static void Tle9210x_SetGenCtrlReg(uint8 u8Group);
/****************************************************************************************
| NAME:    Tle9210x_WriteReg
| CALLED BY:
| PRECONDITIONS:     NA
| INPUT PARAMETERS:    uint16 buf
| RETURN VALUE:     void
| DESCRIPTION:      recovery tle941xy output (overload)
****************************************************************************************/
static void Tle9210x_WriteReg(uint8 u8GroupId,uint8* pu8RegBuf, uint16* pu16WtData)
{
    uint8 l_u8ChipIndex;
    uint8 l_u8ChipNum;
    uint8 l_au8RcvDataBuf[TLE9210X_CHIP_MAX * 3] = {0};
    uint8 l_au8SndDataBuf[TLE9210X_CHIP_MAX * 3] = {0};

    l_u8ChipNum = *cTle9210x_atGroupCfg[u8GroupId].pu8ChipNum;

    if(l_u8ChipNum > 0u)
    {
        /* Reversal data buffer */
        for(l_u8ChipIndex = 0u; l_u8ChipIndex < l_u8ChipNum; l_u8ChipIndex++)
        {
            l_au8SndDataBuf[3*l_u8ChipNum - l_u8ChipIndex] =(uint8)(TLE9210X_BASE_ADDR 
                                            | (uint8)(TLE9210X_LABT_OFF << 7u) 
                                            | (uint8)(pu8RegBuf[(l_u8ChipIndex)] << 1u) 
                                            | TLE9210X_OP_RW_OR_R1C);
            l_au8SndDataBuf[2*l_u8ChipIndex] = pu16WtData[l_u8ChipIndex] ;
            l_au8SndDataBuf[2*l_u8ChipIndex + 1u] = (uint8)(pu16WtData[l_u8ChipIndex] >> 8u);
        }
        /****The last chip control LABT is 1 whether it is daisy chain communication or not********/
        l_au8SndDataBuf[2*l_u8ChipNum + 1u] |= (TLE9210X_LABT_ON << 7u);

        /* Send Read  */
        (void)Spi_SetupEB(cTle9210x_atGroupCfg[u8GroupId].SpiChannel, &l_au8SndDataBuf[0], &l_au8RcvDataBuf[0], (l_u8ChipNum * 3));

        (void)Spi_SyncTransmit(cTle9210x_atGroupCfg[u8GroupId].SpiSequence);

        for(l_u8ChipIndex = 0u; l_u8ChipIndex < l_u8ChipNum; l_u8ChipIndex++)
        {
            sTle9210x_au8GlobalStatus[u8GroupId][l_u8ChipIndex] = l_au8RcvDataBuf[(3u * l_u8ChipNum - l_u8ChipIndex - 1u)];
        }
    }
    else
    {

    }
}

/****************************************************************************************
| NAME:    Tle9210x_ReadReg
| CALLED BY:
| PRECONDITIONS:     NA
| INPUT PARAMETERS:    uint16 buf
| RETURN VALUE:     uint16
| DESCRIPTION:      recovery tle941xy output (overload)
****************************************************************************************/
static void Tle9210x_ReadReg( uint8 u8GroupId,uint8* pu8RegBuf,uint16* pu16ReadBuf)
{

    uint8 l_u8ChipIndex;
    uint8 l_u8ChipNum;
    uint16 l_au8RcvDataBuf[TLE9210X_CHIP_MAX * 3u] = {0};
    uint16 l_au8SndDataBuf[TLE9210X_CHIP_MAX * 3u] = {0};

    l_u8ChipNum = *cTle9210x_atGroupCfg[u8GroupId].pu8ChipNum;

    if(l_u8ChipNum > 0u)
    {
        /* Reversal data buffer */
        for(l_u8ChipIndex = 0u; l_u8ChipIndex < l_u8ChipNum; l_u8ChipIndex++)
        {
            l_au8SndDataBuf[3*l_u8ChipNum - l_u8ChipIndex] =(uint8)(TLE9210X_BASE_ADDR 
                                            | (uint8)(TLE9210X_LABT_OFF << 7u) 
                                            | (uint8)(pu8RegBuf[(l_u8ChipIndex)] << 1u) 
                                            | TLE9210X_OP_READ_ONLY);
        }
        /****The last chip control LABT is 1 whether it is daisy chain communication or not********/
        l_au8SndDataBuf[2*l_u8ChipNum + 1u] |= (uint8)(TLE9210X_LABT_ON << 7u);

        /* Send Read  */
        (void)Spi_SetupEB(cTle9210x_atGroupCfg[u8GroupId].SpiChannel, &l_au8SndDataBuf[0], &l_au8RcvDataBuf[0], (l_u8ChipNum * 2u));

        (void)Spi_SyncTransmit(cTle9210x_atGroupCfg[u8GroupId].SpiSequence);

        for(l_u8ChipIndex = 0u; l_u8ChipIndex < l_u8ChipNum; l_u8ChipIndex++)
        {
            pu16ReadBuf[l_u8ChipIndex] = (uint16)(l_au8RcvDataBuf[l_u8ChipNum * 2u + 1u] << 8u) + l_au8RcvDataBuf[l_u8ChipNum * 2u];
            sTle9210x_au8GlobalStatus[u8GroupId][l_u8ChipIndex] = l_au8RcvDataBuf[(3u*l_u8ChipNum - l_u8ChipIndex - 1u)];
        }
    }
    else
    {

    }

}


static void Tle9210x_SetChipMode(uint8 u8GroupId,uint8 u8Mode)
{
    uint8 j;
    uint8 l_u8ChipNum;

    l_u8ChipNum = *cTle9210x_atGroupCfg[u8GroupId].pu8ChipNum;
    for(j = 0u;j < l_u8ChipNum;j++)
    {
        switch(u8Mode)
        {
            case TLE9210X_MODE_SLEEP:
                {
                    Dio_WriteChannel(cTle9210x_atChipCfg[u8GroupId][j].u8EnPinCtrl, STD_OFF);
                    sTle9210x_au8ChipMode[u8GroupId][j] = TLE9210X_MODE_SLEEP;
                }
                break;
            case TLE9210X_MODE_NORMAL:
                Dio_WriteChannel(cTle9210x_atChipCfg[u8GroupId][j].u8EnPinCtrl, STD_ON);
                sTle9210x_au8ChipMode[u8GroupId][j] = TLE9210X_MODE_NORMAL;
                break;
            case TLE9210X_MODE_FAIL_SAFE:
                Dio_WriteChannel(cTle9210x_atChipCfg[u8GroupId][j].u8EnPinCtrl, STD_ON);
                sTle9210x_au8ChipMode[u8GroupId][j] = TLE9210X_MODE_FAIL_SAFE;
                break;
            default:
                Dio_WriteChannel(cTle9210x_atChipCfg[u8GroupId][j].u8EnPinCtrl, STD_ON);
                sTle9210x_au8ChipMode[u8GroupId][j] = TLE9210X_MODE_NORMAL;
                break;
        }
    }
}

static void Tle9210x_GetChipMode(uint8 u8GroupId,uint8 u8ChipId,uint8* pu8Mode)
{
    *pu8Mode = sTle9210x_au8ChipMode[u8GroupId][u8ChipId];
}

static void Tle9210x_SetGenCtrlReg(uint8 u8Group)
{

    uint8 j;
    uint8 l_au8RegBuf[TLE9210X_CHIP_MAX] = {0};
    uint8 l_au16DataBuf[TLE9210X_CHIP_MAX] = {0};
    uint8 l_u8ChipNum;

    l_u8ChipNum = *cTle9210x_atGroupCfg[u8Group].pu8ChipNum;

    for(j = 0u;j < l_u8ChipNum;j++)
    {
        l_au8RegBuf[j] = TLE9210X_GENCTRL1;
    }
    Tle9210x_ReadReg(u8Group,l_au8RegBuf,l_au16DataBuf);

    for(j = 0u;j < l_u8ChipNum;j++)
    {
        l_au16DataBuf[j] &= 0xfdfdu;
        l_au16DataBuf[j] |= ((uint16)(cTle9210x_atChipCfg[u8Group][j].REG_BANK << 9u)
                        | (uint16)(cTle9210x_atChipCfg[u8Group][j].WDPER << 1u));
        sTle9210x_au16GenCtrl1[u8Group][j] = l_au16DataBuf[j];
        sTle9210x_abREGBANKSts[u8Group][j] = cTle9210x_atChipCfg[u8Group][j].REG_BANK;
    }
    Tle9210x_WriteReg(u8Group,l_au8RegBuf,l_au16DataBuf);

    for(j = 0u;j < l_u8ChipNum;j++)
    {
        l_au8RegBuf[j] = TLE9210X_GENCTRL2;
    }
    Tle9210x_ReadReg(u8Group,l_au8RegBuf,l_au16DataBuf);

    for(j = 0u;j < l_u8ChipNum;j++)
    {
        l_au16DataBuf[j] &= 0xfdffu;
        l_au16DataBuf[j] |= (uint16)(cTle9210x_atChipCfg[u8Group][j].WDDIS << 9u);
    }
    Tle9210x_WriteReg(u8Group,l_au8RegBuf,l_au16DataBuf);

}

static void Tle9210x_SetVDSReg(uint8 u8Group)
{
    uint8 j;
    uint8 l_au8RegBuf[TLE9210X_CHIP_MAX] = {0};
    uint8 l_au16DataBuf[TLE9210X_CHIP_MAX] = {0};
    uint8 l_u8ChipNum;

    l_u8ChipNum = *cTle9210x_atGroupCfg[u8Group].pu8ChipNum;

    for(j = 0u;j < l_u8ChipNum;j++)
    {
        l_au8RegBuf[j] = TLE9210X_VDS1;
        l_au16DataBuf[j] |= ((cTle9210x_atHbChnCfg[u8Group][j][0].u8HBDrainSrcOv)
                |(uint16)(cTle9210x_atHbChnCfg[u8Group][j][1].u8HBDrainSrcOv << 3u)
                |(uint16)(cTle9210x_atHbChnCfg[u8Group][j][2].u8HBDrainSrcOv << 6u)
                |(uint16)(cTle9210x_atHbChnCfg[u8Group][j][3].u8HBDrainSrcOv << 9u)
                |(uint16)(cTle9210x_atHbChnCfg[u8Group][j][0].bHBDrainSrcMonit << 12u)
                |(uint16)(cTle9210x_atHbChnCfg[u8Group][j][0].bHBDrainSrcMonit << 13u)
                |(uint16)(cTle9210x_atHbChnCfg[u8Group][j][0].bHBDrainSrcMonit << 14u)
                |(uint16)(cTle9210x_atHbChnCfg[u8Group][j][0].bHBDrainSrcMonit << 15u));
    }

    Tle9210x_WriteReg(u8Group,l_au8RegBuf,l_au16DataBuf);

#if(TLE9210X_TLE92108_CHIP_EN == STD_ON)
    for(j = 0u;j < l_u8ChipNum;j++)
    {
        l_au8RegBuf[j] = TLE9210X_VDS2;

        l_au16DataBuf[j] |= ((cTle9210x_atHbChnCfg[u8Group][j][4].u8HBDrainSrcOv)
        |(uint16)(cTle9210x_atHbChnCfg[u8Group][j][5].u8HBDrainSrcOv << 3u)
        |(uint16)(cTle9210x_atHbChnCfg[u8Group][j][6].u8HBDrainSrcOv << 6u)
        |(uint16)(cTle9210x_atHbChnCfg[u8Group][j][7].u8HBDrainSrcOv << 9u)
        |(uint16)(cTle9210x_atHbChnCfg[u8Group][j][4].bHBDrainSrcMonit << 12u)
        |(uint16)(cTle9210x_atHbChnCfg[u8Group][j][5].bHBDrainSrcMonit << 13u)
        |(uint16)(cTle9210x_atHbChnCfg[u8Group][j][6].bHBDrainSrcMonit << 14u)
        |(uint16)(cTle9210x_atHbChnCfg[u8Group][j][7].bHBDrainSrcMonit << 15u));
    }
    Tle9210x_WriteReg(u8Group,l_au8RegBuf,l_au16DataBuf);
#endif

}

static void Tle9210x_SetHbOutputReg(uint8 u8Group)
{

    uint8 j;
    uint8 l_au8RegBuf[TLE9210X_CHIP_MAX] = {0};
    uint8 l_au16DataBuf[TLE9210X_CHIP_MAX] = {0};
    uint8 l_u8ChipNum;

    l_u8ChipNum = *cTle9210x_atGroupCfg[u8Group].pu8ChipNum;
    /***OUT1-OUT4**/
    for(j = 0u;j < l_u8ChipNum;j++)
    {
        l_au8RegBuf[j] = TLE9210X_HBMODE;
        l_au16DataBuf[j] = (sTle9210x_au8HbOutSts[u8Group][j][0] 
                        | (uint8)(sTle9210x_au8HbOutSts[u8Group][j][1] << 2u)
                        | (uint8)(sTle9210x_au8HbOutSts[u8Group][j][2] << 4u)
                        | (uint8)(sTle9210x_au8HbOutSts[u8Group][j][3] << 6u)
                        | (uint8)(sTle9210x_au8HbOutSts[u8Group][j][4] << 8u)
                        | (uint8)(sTle9210x_au8HbOutSts[u8Group][j][5] << 10u)
                        | (uint8)(sTle9210x_au8HbOutSts[u8Group][j][6] << 12u)
                        | (uint8)(sTle9210x_au8HbOutSts[u8Group][j][7] << 14u));
    }
    Tle9210x_WriteReg(u8Group,l_au8RegBuf,l_au16DataBuf);

}

static void Tle9210x_SetPwmActOrFw(uint8 u8Group)
{
    uint8 j;
    uint8 l_au8RegBuf[TLE9210X_CHIP_MAX] = {0};
    uint8 l_au16DataBuf[TLE9210X_CHIP_MAX] = {0};
    uint8 l_u8ChipNum;

    l_u8ChipNum = *cTle9210x_atGroupCfg[u8Group].pu8ChipNum;
    for(j = 0u;j < l_u8ChipNum;j++)
    {
        if(sTle9210x_abREGBANKSts[u8Group][j] == TLE9210X_REG_BANK_OFF)
        {
            l_au8RegBuf[j] = TLE9210X_CCP_BLK2_ACT;
        }
        else
        {
            l_au8RegBuf[j] = TLE9210X_CCP_BLK2_FW;
        }
    }
    Tle9210x_WriteReg(u8Group,l_au8RegBuf,l_au16DataBuf);

    for(j = 0u;j < l_u8ChipNum;j++)
    {
        if(sTle9210x_abREGBANKSts[u8Group][j] == TLE9210X_REG_BANK_OFF)
        {
            l_au8RegBuf[j] = TLE9210X_ST_ICHG;
        }
        else
        {
            l_au8RegBuf[j] = TLE9210X_PWM_PCHG_INIT;
        }
    }
    Tle9210x_WriteReg(u8Group,l_au8RegBuf,l_au16DataBuf);

    for(j = 0u;j < l_u8ChipNum;j++)
    {
        if(sTle9210x_abREGBANKSts[u8Group][j] == TLE9210X_REG_BANK_OFF)
        {
            l_au8RegBuf[j] = TLE9210X_PWM_ICHG_ACT;
        }
        else
        {
            l_au8RegBuf[j] = TLE9210X_PWM_ICHG_FW;
        }
    }
    Tle9210x_WriteReg(u8Group,l_au8RegBuf,l_au16DataBuf);

    for(j = 0u;j < l_u8ChipNum;j++)
    {
        if(sTle9210x_abREGBANKSts[u8Group][j] == TLE9210X_REG_BANK_OFF)
        {
            l_au8RegBuf[j] = TLE9210X_PWM_IDCHG_ACT;
        }
        else
        {
            l_au8RegBuf[j] = TLE9210X_PWM_PDCHG_INIT;
        }
    }
    Tle9210x_WriteReg(u8Group,l_au8RegBuf,l_au16DataBuf);

    for(j = 0u;j < l_u8ChipNum;j++)
    {
        if(sTle9210x_abREGBANKSts[u8Group][j] == TLE9210X_REG_BANK_OFF)
        {
            l_au8RegBuf[j] = TLE9210X_PWM_ICHGMAX_CCP_BLK3_ACT;
        }
        else
        {
            l_au8RegBuf[j] = TLE9210X_PWM_ICHGMAX_CCP_BLK3_FW;
        }
    }
    Tle9210x_WriteReg(u8Group,l_au8RegBuf,l_au16DataBuf);
}

static void Tle9210x_SetRegBank(uint8 u8Group,boolean bEn)
{
    uint8 j;
    uint8 l_au8RegBuf[TLE9210X_CHIP_MAX] = {0};
    uint8 l_au16DataBuf[TLE9210X_CHIP_MAX] = {0};
    uint8 l_u8ChipNum;

    l_u8ChipNum = *cTle9210x_atGroupCfg[u8Group].pu8ChipNum;

    for(j = 0u;j < l_u8ChipNum;j++)
    {
        l_au8RegBuf[j] = TLE9210X_GENCTRL1;
    }
    Tle9210x_ReadReg(u8Group,l_au8RegBuf,l_au16DataBuf);
    for(j = 0u;j < l_u8ChipNum;j++)
    {
        if(TRUE == bEn)
        {
            SETBIT_U16(l_au16DataBuf[j],9u);
            sTle9210x_abREGBANKSts[u8Group][j] = TLE9210X_REG_BANK_ON;
        }
        else
        {
            CLRBIT_U16(l_au16DataBuf[j],9u);
            sTle9210x_abREGBANKSts[u8Group][j] = TLE9210X_REG_BANK_OFF;
        }
    }

    Tle9210x_WriteReg(u8Group,l_au8RegBuf,l_au16DataBuf);
}

static void Tle9210x_SetPwmMappingReg(uint8 u8Group)
{
    uint8 j;
    uint8 l_au8RegBuf[TLE9210X_CHIP_MAX] = {0};
    uint8 l_au16DataBuf[TLE9210X_CHIP_MAX] = {0};
    uint8 l_u8ChipNum;
    Tle9210x_SetRegBank(u8Group,TLE9210X_REG_BANK_OFF);
    l_u8ChipNum = *cTle9210x_atGroupCfg[u8Group].pu8ChipNum;
    /***OUT1-OUT4**/
    for(j = 0u;j < l_u8ChipNum;j++)
    {
        l_au8RegBuf[j] = TLE9210X_PWMSET;
        l_au16DataBuf[j] = (cTle9210x_atPwmChnCfg[u8Group][j][0].bPwmEn
                        | (uint16)(cTle9210x_atPwmChnCfg[u8Group][j][0].u8PwmMapChn << 1u)
                        | (uint16)(cTle9210x_atPwmChnCfg[u8Group][j][1].bPwmEn << 4u)
                        | (uint16)(cTle9210x_atPwmChnCfg[u8Group][j][1].u8PwmMapChn << 5u)
                        | (uint16)(cTle9210x_atPwmChnCfg[u8Group][j][2].bPwmEn << 8u)
                        | (uint16)(cTle9210x_atPwmChnCfg[u8Group][j][2].u8PwmMapChn << 9u));
    }
    Tle9210x_WriteReg(u8Group,l_au8RegBuf,l_au16DataBuf);

}

static void Tle9210x_SetPwmDelayTimeReg(uint8 u8Group)
{
    uint8 j;
    uint8 l_au8RegBuf[TLE9210X_CHIP_MAX] = {0};
    uint8 l_au16DataBuf[TLE9210X_CHIP_MAX] = {0};
    uint8 l_u8ChipNum;
    uint8 l_u8ErrCnt;
    uint8 l_u8RetVal;

    l_u8ChipNum = *cTle9210x_atGroupCfg[u8Group].pu8ChipNum;
    /***PWM1**/
    for(j = 0u;j < l_u8ChipNum;j++)
    {
        l_au8RegBuf[j] = TLE9210X_TDON_OFF1;
        l_au16DataBuf[j] = (cTle9210x_atPwmChnCfg[u8Group][j][0].u8TurnOnTime
                        | (uint16)(cTle9210x_atPwmChnCfg[u8Group][j][0].u8TurnOffTime << 8u));
    }
    Tle9210x_WriteReg(u8Group,l_au8RegBuf,l_au16DataBuf);

    /***PWM2**/
    for(j = 0u;j < l_u8ChipNum;j++)
    {
        l_au8RegBuf[j] = TLE9210X_TDON_OFF2;
        l_au16DataBuf[j] = (cTle9210x_atPwmChnCfg[u8Group][j][1].u8TurnOnTime
                        | (uint16)(cTle9210x_atPwmChnCfg[u8Group][j][1].u8TurnOffTime << 8u));
    }
    Tle9210x_WriteReg(u8Group,l_au8RegBuf,l_au16DataBuf);

    /***PWM3**/
    for(j = 0u;j < l_u8ChipNum;j++)
    {
        l_au8RegBuf[j] = TLE9210X_TDON_OFF3;
        l_au16DataBuf[j] = (cTle9210x_atPwmChnCfg[u8Group][j][2].u8TurnOnTime
                        | (uint16)(cTle9210x_atPwmChnCfg[u8Group][j][2].u8TurnOffTime << 8u));
    }
    Tle9210x_WriteReg(u8Group,l_au8RegBuf,l_au16DataBuf);

}

static void Tle9210x_GetAllGenSts(uint8 u8Group)
{
    uint8 j;
    uint8 l_au8RegBuf[TLE9210X_CHIP_MAX] = {0};
    uint8 l_au16DataBuf[TLE9210X_CHIP_MAX] = {0};
    uint8 l_u8ChipNum;

    l_u8ChipNum = *cTle9210x_atGroupCfg[u8Group].pu8ChipNum;

    for(j = 0u;j < l_u8ChipNum;j++)
    {
        l_au8RegBuf[j] = TLE9210X_GENSTAT;
    }
    Tle9210x_ReadReg(u8Group,l_au8RegBuf,l_au16DataBuf);
    for(j = 0u;j < l_u8ChipNum;j++)
    {
        sTle9210x_atGenStsReport[u8Group][j].u16GENSTAT = l_au16DataBuf[j];
    }

    for(j = 0u;j < l_u8ChipNum;j++)
    {
        l_au8RegBuf[j] = TLE9210X_DSOV;
    }
    Tle9210x_ReadReg(u8Group,l_au8RegBuf,l_au16DataBuf);
    for(j = 0u;j < l_u8ChipNum;j++)
    {
        sTle9210x_atGenStsReport[u8Group][j].u16DSOV = l_au16DataBuf[j];
    }

    for(j = 0u;j < l_u8ChipNum;j++)
    {
        l_au8RegBuf[j] = TLE9210X_HBVOUT_PWMERR;
    }
    Tle9210x_ReadReg(u8Group,l_au8RegBuf,l_au16DataBuf);
    for(j = 0u;j < l_u8ChipNum;j++)
    {
        sTle9210x_atGenStsReport[u8Group][j].u16HBVOUT_PWMERR = l_au16DataBuf[j];
    }

    for(j = 0u;j < l_u8ChipNum;j++)
    {
        l_au8RegBuf[j] = TLE9210X_EFF_TDON_OFF1;
    }
    Tle9210x_ReadReg(u8Group,l_au8RegBuf,l_au16DataBuf);
    for(j = 0u;j < l_u8ChipNum;j++)
    {
        sTle9210x_atGenStsReport[u8Group][j].u16EFF_TDON_OFF1 = l_au16DataBuf[j];
    }

    for(j = 0u;j < l_u8ChipNum;j++)
    {
        l_au8RegBuf[j] = TLE9210X_EFF_TDON_OFF2;
    }
    Tle9210x_ReadReg(u8Group,l_au8RegBuf,l_au16DataBuf);
    for(j = 0u;j < l_u8ChipNum;j++)
    {
        sTle9210x_atGenStsReport[u8Group][j].u16EFF_TDON_OFF2 = l_au16DataBuf[j];
    }

    for(j = 0u;j < l_u8ChipNum;j++)
    {
        l_au8RegBuf[j] = TLE9210X_EFF_TDON_OFF3;
    }
    Tle9210x_ReadReg(u8Group,l_au8RegBuf,l_au16DataBuf);
    for(j = 0u;j < l_u8ChipNum;j++)
    {
        sTle9210x_atGenStsReport[u8Group][j].u16EFF_TDON_OFF3 = l_au16DataBuf[j];
    }

    for(j = 0u;j < l_u8ChipNum;j++)
    {
        l_au8RegBuf[j] = TLE9210X_TRISE_FALL1;
    }
    Tle9210x_ReadReg(u8Group,l_au8RegBuf,l_au16DataBuf);
    for(j = 0u;j < l_u8ChipNum;j++)
    {
        sTle9210x_atGenStsReport[u8Group][j].u16TRISE_FALL1 = l_au16DataBuf[j];
    }

    for(j = 0u;j < l_u8ChipNum;j++)
    {
        l_au8RegBuf[j] = TLE9210X_TRISE_FALL2;
    }
    Tle9210x_ReadReg(u8Group,l_au8RegBuf,l_au16DataBuf);
    for(j = 0u;j < l_u8ChipNum;j++)
    {
        sTle9210x_atGenStsReport[u8Group][j].u16TRISE_FALL2 = l_au16DataBuf[j];
    }

    for(j = 0u;j < l_u8ChipNum;j++)
    {
        l_au8RegBuf[j] = TLE9210X_TRISE_FALL3;
    }
    Tle9210x_ReadReg(u8Group,l_au8RegBuf,l_au16DataBuf);
    for(j = 0u;j < l_u8ChipNum;j++)
    {
        sTle9210x_atGenStsReport[u8Group][j].u16TRISE_FALL3 = l_au16DataBuf[j];
    }

    for(j = 0u;j < l_u8ChipNum;j++)
    {
        l_au8RegBuf[j] = TLE9210X_DEVID;
    }
    Tle9210x_ReadReg(u8Group,l_au8RegBuf,l_au16DataBuf);
    for(j = 0u;j < l_u8ChipNum;j++)
    {
        sTle9210x_atGenStsReport[u8Group][j].u16DEVID = l_au16DataBuf[j];
    }
}

static void Tle9210x_OVDiagnostic(uint8 u8Group)
{
    uint8 j;
    uint8 k;
    uint8 l_u8Chn;
    uint8 l_au8RegBuf[TLE9210X_CHIP_MAX] = {0};
    uint8 l_au16DataBuf[TLE9210X_CHIP_MAX] = {0};
    uint8 l_u8ChipNum;
    uint8 l_u8ErrCnt;

    for(j = 0u;j < l_u8ChipNum;j++)
    {
        l_au8RegBuf[j] = TLE9210X_DSOV;
    }
    Tle9210x_ReadReg(u8Group,l_au8RegBuf,l_au16DataBuf);
    for(j = 0u;j < l_u8ChipNum;j++)
    {
        sTle9210x_atGenStsReport[u8Group][j].u16DSOV = l_au16DataBuf[j];
        for(k = 0u;k < 16u;k + 2u)
        {
            l_u8Chn = (uint8)(k/2u);
            sTle9210x_atDiagResult[u8Group][j][l_u8Chn].eShort2Vcc = 
                (TRUE == (GETBIT_U16(sTle9210x_atGenStsReport[u8Group][j].u16DSOV,k)
                ||GETBIT_U16(sTle9210x_atGenStsReport[u8Group][j].u16DSOV,(k+1u))))
                ? PFM_DDS_POS : PFM_DDS_NEG;
        }
        if(l_au16DataBuf[j] > 0u)
        {
            l_u8ErrCnt++;
        }
        else
        {
            /*Nothing to do*/
        }
    }
    if(l_u8ErrCnt > 0u)
    {
        memset(l_au16DataBuf,0u,sizeof(l_au16DataBuf));
        Tle9210x_WriteReg(u8Group,l_au8RegBuf,l_au16DataBuf);
    }
}

static void Tle9210x_SetPwmDutyOut(uint8 u8Group)
{
    uint8 j;
    uint8 k;
    uint8 l_u8ChipNum;

    l_u8ChipNum = *cTle9210x_atGroupCfg[u8Group].pu8ChipNum;
    /***PWM1 - PWM3**/
    for(j = 0u;j < l_u8ChipNum;j++)
    {
        for(k = 0u;k < 3u; k++)
        {
            if(cTle9210x_atPwmChnCfg[u8Group][j][0].bPwmEn == TLE9210X_PWM_ENABLE)
            {
                Pwm_SetDutyCycle(cTle9210x_atPwmChnCfg[u8Group][j][0].u8PwmChn,
                    sTle9210x_au8PwmDuty[u8Group][j][0]);
            }
        }
    }
}

void Tle9210x_Init(void)
{
    uint8 i;

    memset(sTle9210x_au8HbOutSts,0u,sizeof(sTle9210x_au8HbOutSts));

    for(i = 0u;i < TLE9210X_GROUP_MAX;i++)
    {
        Tle9210x_SetChipMode(i,TLE9210X_MODE_NORMAL);
        Tle9210x_SetGenCtrlReg(i);
        Tle9210x_SetPwmMappingReg(i);
        Tle9210x_SetPwmDelayTimeReg(i);
        Tle9210x_GetAllGenSts(i);
        Tle9210x_SetVDSReg(i);

        Tle9210x_SetHbOutputReg(i);
    }
}


void Tle9210x_MainFunction(void)
{
    uint8 i;

    for(i = 0u;i < TLE9210X_GROUP_MAX;i++)
    {
        Tle9210x_OVDiagnostic(i);
        Tle9210x_SetHbOutputReg(i);
        Tle9210x_SetPwmDutyOut(i);
    }
}

void Tle9210x_DeInit(void)
{
    uint8 i;

    memset(sTle9210x_au8HbOutSts,0u,sizeof(sTle9210x_au8HbOutSts));
    memset(sTle9210x_au8PwmDuty,0u,sizeof(sTle9210x_au8PwmDuty));
    for(i = 0u;i < TLE9210X_GROUP_MAX;i++)
    {
        Tle9210x_SetHbOutputReg(i);
        Tle9210x_SetPwmDutyOut(i);
    }
}

void Tle9210x_WriteHbChn(uint8 u8GroupId, uint8 u8ChipId,uint8 u8ChnId, uint8 u8Val)
{
    if((u8GroupId < (uint8)TLE9210X_GROUP_MAX)
    &&(u8ChipId < (uint8)TLE9210X_CHIP_MAX)
    &&(u8ChnId < (uint8)TLE9210X_HB_CHN_MAX))
    {
        sTle9210x_au8HbOutSts[u8GroupId][u8ChipId][u8ChnId] = u8Val;
    }
}

void Tle9210x_WritePwmChn(uint8 u8GroupId, uint8 u8ChipId,uint8 u8PwmChn, uint8 u8Val)
{
    if((u8GroupId < (uint8)TLE9210X_GROUP_MAX)
    &&(u8ChipId < (uint8)TLE9210X_CHIP_MAX)
    &&(u8PwmChn < (uint8)TLE9210X_PWM_CHN_MAX))
    {
        sTle9210x_au8PwmDuty[u8GroupId][u8ChipId][u8PwmChn] = u8Val;
    }
}

void Tle9210x_TriggerWdg(uint8 u8Group)
{

    uint8 j;
    uint8 l_au8RegBuf[TLE9210X_CHIP_MAX] = {0};
    uint8 l_au16DataBuf[TLE9210X_CHIP_MAX] = {0};
    uint8 l_u8ChipNum;

    l_u8ChipNum = *cTle9210x_atGroupCfg[u8Group].pu8ChipNum;

    for(j = 0u;j < l_u8ChipNum;j++)
    {
        l_au8RegBuf[j] = TLE9210X_GENCTRL1;

        if((sTle9210x_au16GenCtrl1[u8Group][j]  & 0x0001) == 0x0001u)
        {
            sTle9210x_au16GenCtrl1[u8Group][j] &= 0xfffeu;
        }
        else
        {
            sTle9210x_au16GenCtrl1[u8Group][j] |= 0x0001u;
        }
        l_au16DataBuf[j] = sTle9210x_au16GenCtrl1[u8Group][j];
    }

    Tle9210x_WriteReg(u8Group,l_au8RegBuf,l_au16DataBuf);
}
