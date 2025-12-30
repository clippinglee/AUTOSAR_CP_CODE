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
#ifndef _TLE941XY_H_
#define _TLE941XY_H_

#include "Tle941xy_Types.h"
#include "Tle941xy_HwCfg.h"

extern void Tle941xy_Init(void);
extern void Tle941xy_MainFunction(void);
extern void Tle941xy_DeInit(void);
extern void Tle941xy_WriteOhbChn(uint8 u8GroupId, uint8 u8ChipId,uint8 u8ChnId, uint8 u8Val);
extern void Tle941xy_WritePwmChn(uint8 u8GroupId, uint8 u8ChipId,uint8 u8PwmChn, uint8 u8Val);

#endif
