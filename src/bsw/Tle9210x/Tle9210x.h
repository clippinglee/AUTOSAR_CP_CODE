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
#ifndef _TLE9210X_H_
#define _TLE9210X_H_
#include "Tle9210x_HwCfg.h"
#include "Tle9210x_Types.h"


extern void Tle9210x_Init(void);
extern void Tle9210x_MainFunction(void);
extern void Tle9210x_DeInit(void);
extern void Tle9210x_WriteOhbChn(uint8 u8GroupId, uint8 u8ChipId,uint8 u8ChnId, uint8 u8Val);
extern void Tle9210x_WritePwmChn(uint8 u8GroupId, uint8 u8ChipId,uint8 u8PwmChn, uint8 u8Val);

#endif
