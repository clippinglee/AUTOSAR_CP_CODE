/*****************************************************************************************************************
******************************************************************************************************************
*  Copyright (C) .                                                                            
*  All rights reserved.                                                                                           
******************************************************************************************************************
*  FileName: BJT                                                                                             
*  Content:  BJT family drive
*  Category: 
******************************************************************************************************************
*  Revision Management                                                                                            
*  yyyy.mm.dd    name              version      description                                                       
*  ----------    --------          -------      -----------------------------------                               
*  2022.04.18    clipping            v0001        Frist edit                                                        
******************************************************************************************************************
******************************************************************************************************************/
/* Include Headerfiles  */
#include "Bjt_HwCfg.h"
#include "AdcIf.h"
const Bjt_ChnCfgType cBjt_atChannelInputCfg[BJT_ID_MAX] = 
{
    {BJT_ID_0,BJT_DIO,DioConf_DioChannel_DioChannel_P31_11,BJT_DIO_PWM_INVALIDVAL,0xFFF,0xFFF},
    {BJT_ID_1,BJT_DIO,DioConf_DioChannel_DioChannel_P31_12,BJT_DIO_PWM_INVALIDVAL,0xFFF,0xFFF},

};

const uint8 cBjt_au8AdcEidMap[BJT_ID_MAX] =
{
    1,2
};

