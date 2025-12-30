/*****************************************************************************************************************
******************************************************************************************************************
*  Copyright (C) .                                                                            
*  All rights reserved.                                                                                           
******************************************************************************************************************
*  FileName: Bjt                                                                                             
*  Content:  Bjt family drive
*  Category:                                                                              				          
******************************************************************************************************************
*  Revision Management                                                                                            
*  yyyy.mm.dd    name              version      description                                                       
*  ----------    --------          -------      -----------------------------------                               
*  2022.04.18    clipping            V0001        Frist edit                                                        
******************************************************************************************************************
******************************************************************************************************************/
/* Include Headerfiles  */
#include "Bjt.h"
#include "Bjt_Types.h"
#include "Bjt_HwCfg.h"

#include "AdcIf.h"
#include "PFM.h"
#include <string.h>
#include "LiBool.h"
/* PRQA S 0314 EOF*/

/*******************************************************************************
**  Local  variable definitions
*******************************************************************************/

/* the fault status of all channels*/
static PFM_DefectReportState_t sBjt_atDiagResult[BJT_ID_MAX];

/*******************************************************************************
**  Global  variable definitions
*******************************************************************************/
static uint16 sBjt_au16PwmOutDuty[BJT_ID_MAX];
static boolean sBjt_abDoValue[BJT_ID_MAX];
/* the record of ON/OFF status of all channels */
static uint32 sBjt_u32ChnSts;
/* ADC value of feedback diagnostic signals of all channels */
static uint16 gBjt_au16DiagAdcV[BJT_ID_MAX];

#define BJT_GETCHANSTATE(port)    GETBIT_U32(sBjt_u32ChnSts, port)
/*******************************************************************************
**  Local  Function definitions
*******************************************************************************/
static void Bjt_GetDiagAdVal(void);
static void Bjt_DiagHandle(void);
static boolean Bjt_JudgePwmDuty(uint16 u16NominalValue);
static void Bjt_WriteOutput(void);
/*******************************************************************************
**  Global  Function definitions
*******************************************************************************/

/****************************************************************
 process: Bjt_vDiagHandle
 purpose: The diagnostic handler for BJT module. Invoked every 10ms.
          Figure out if openload, short to battery or short to ground
          fault has occured. Always report current fault status to 
          PFM module. Debounce is performed in PFM module.
 ****************************************************************/
static void Bjt_DiagHandle(void)
{
    boolean l_bChanState;
    uint8   l_u8Port;
    PFM_PhysicalId_e l_eFid; 
    
    /* Go through all channels and perform diagnostic operation.
       Report diagnosing result to PFM. */
    for( l_u8Port = 0u; l_u8Port < (uint8)BJT_ID_MAX; l_u8Port ++ )
    {
        /* Get the index of current channel in PFM module*/
        l_bChanState = (boolean)(BJT_GETCHANSTATE(l_u8Port) ? TRUE : FALSE);
        /* if diagnosing channel selection equals this channel (channel 0 or channel 1),
           which means the ADC sample value belongs to this channel, diagnosing can be 
           performed. Otherwise the diagnosing of this channel should wait until next cycle*/
        if( (GETBIT_U32(sBjt_u32ChnSts,l_u8Port)
            || (l_bChanState == (boolean)FALSE)))
        {
            if(gBjt_au16DiagAdcV[l_u8Port] <= cBjt_atChannelInputCfg[l_u8Port].u16OLDiagAdcVal)
            {
                sBjt_atDiagResult[l_u8Port].eOpenLoad  = PFM_DDS_POS;
            }
            else if(gBjt_au16DiagAdcV[l_u8Port] >= cBjt_atChannelInputCfg[l_u8Port].u16ShortDiagAdcVal)
            {
                sBjt_atDiagResult[l_u8Port].eShort2Gnd  = PFM_DDS_POS;
            }
            else
            {
                sBjt_atDiagResult[l_u8Port].eOpenLoad  = PFM_DDS_ING;
                sBjt_atDiagResult[l_u8Port].eShort2Vcc = PFM_DDS_ING;
            }
            sBjt_atDiagResult[l_u8Port].eShort2Gnd = PFM_DDS_ING;
        }
        else   /* If this channel is not selected as feedback source, wait for next cycle */
        {
            sBjt_atDiagResult[l_u8Port].eOpenLoad  = PFM_DDS_ING;
            sBjt_atDiagResult[l_u8Port].eShort2Vcc = PFM_DDS_ING;
            sBjt_atDiagResult[l_u8Port].eShort2Gnd = PFM_DDS_ING;
        }
        Pfm_DefectReport(l_eFid, sBjt_atDiagResult[l_u8Port].eOpenLoad, sBjt_atDiagResult[l_u8Port].eShort2Vcc, sBjt_atDiagResult[l_u8Port].eShort2Gnd);
    }
}

/****************************************************************
 process: Bjt_vGetDiagAdVal
 purpose: This function records diagnostic raw values and external 
          control signal from ADC buffer.
          Some lamps have redundant control which is realized by a 
          hardwire. If BJT channel is off while external control 
          is on, there is feedback current on diagnostic pin anyway.
          This function recognises this situation by reading ADC
          value of corresponding hardwire control.
 ****************************************************************/
static void Bjt_GetDiagAdVal(void)
{
    uint8 l_u8Port;
    uint8 l_u8DiagChan;

    for (l_u8Port = 0u; l_u8Port < (uint8)BJT_ID_MAX; l_u8Port ++)
    {
        /* Get the AD value of diagnostic feedback */
        gBjt_au16DiagAdcV[l_u8Port] = AdcIf_GetAdcValue(cBjt_au8AdcEidMap[l_u8Port]);

    }
}

/****************************************************************
 process: Bjt_vTurnOffAll
 purpose: Turn off all output channel ae weel as re-initilizing
          the status record "gBjt_u32ChanStat".
 ****************************************************************/
void Bjt_TurnOffAll(void)
{
    uint8 l_u8Port;
    for(l_u8Port = 0u; l_u8Port < (uint8)BJT_ID_MAX; l_u8Port++)
    {
        Bjt_WriteDoChn(l_u8Port, 0u);
    }
    /* update all channel state record to off */
    sBjt_u32ChnSts = 0u;   
    Bjt_WriteOutput();
}

static void Bjt_WriteOutput(void)
{
    uint8 i;
    for(i = 0;i < (uint8)BJT_ID_MAX;i++)
    {
        if(BJT_PWM == cBjt_atChannelInputCfg[i].eBjt_Type)
        {
            Pwm_SetDutyCycle(cBjt_atChannelInputCfg[i].u8BjtPwmCntrl, sBjt_au16PwmOutDuty[i]);
        }
        else if( BJT_DIO== cBjt_atChannelInputCfg[i].eBjt_Type)
        {
            Dio_WriteChannel(cBjt_atChannelInputCfg[i].u8BjtDioCntrl, sBjt_abDoValue[i]);
        }
        else
        {
            /*do nothing*/
        }
    }

}

/****************************************************************
 process: Bjt_v10ms
 purpose: BJT module main function. invuked every 10 ms.
 ****************************************************************/
void Bjt_MainFunction(void)
{
    Bjt_GetDiagAdVal();
    Bjt_DiagHandle();
    Bjt_WriteOutput();
}

/****************************************************************
 process: Bjt_vInit
 purpose: Init process.
 ****************************************************************/
void Bjt_Init(void)
{
    /* initialize the global diagnostic variables */
    (void)memset((void *)gBjt_au16DiagAdcV, 0, sizeof(uint16) * (uint8)BJT_ID_MAX);                /* initialize AD values of all diagnostic channels */
    (void)memset((void *)sBjt_atDiagResult,0,sizeof(PFM_DefectReportState_t) * (uint8)BJT_ID_MAX); /* all diagnostic result = PFM_DDS_ING(0) */

    /* turn off all outputs and initialize state record to all-off */
    Bjt_TurnOffAll();  
}
/****************************************************************
 process: Bjt_vDeInit
 purpose: Put all BJT chips to initialized state as well as 
          turning off all outputs.
          If application layer does not update BJT on/off status right 
          after waking up, the outputs should not be closed here.
 ****************************************************************/
void Bjt_DeInit(void)
{  
    /* turn off all outputs and initialize state record to all-off */
    Bjt_TurnOffAll();
    /* initialize the global diagnostic variables */
    (void)memset((void *)gBjt_au16DiagAdcV, 0, sizeof(uint16) * (uint8)BJT_ID_MAX);                  /* zero all diagnostic signal AD value */    (void)memset((void *)sBjt_atDiagResult, 0, sizeof(PFM_DefectReportState_t)*  (uint8)BJT_ID_MAX); /* all diagnostic result = PFM_DDS_ING(0) */

}

/****************************************************************
 process: Bjt_vWriteDoChn
 purpose: Write Channel Value.
 ****************************************************************/
void Bjt_WriteDoChn(uint8 u8Chn, uint16 u16Val)
{
    if(BJT_PWM == cBjt_atChannelInputCfg[u8Chn].eBjt_Type)
    {
        sBjt_au16PwmOutDuty[u8Chn] = u16Val;
    }
    else if( BJT_DIO== cBjt_atChannelInputCfg[u8Chn].eBjt_Type)
    {
        sBjt_abDoValue[u8Chn] = (boolean)u16Val;
    }
    else
    {
        /*do nothing*/
    }

    if (u16Val > 0u)
    {
        sBjt_u32ChnSts |= (uint32)1u << u8Chn;
    }
    else
    {
        sBjt_u32ChnSts &= 0xFFFFFFFFul - ((uint32)1u << u8Chn);
    }
}
