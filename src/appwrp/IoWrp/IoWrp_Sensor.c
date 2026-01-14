/*****************************************************************************************************************
******************************************************************************************************************
*  Copyright (C) .
*  All rights reserved.
******************************************************************************************************************
*  FileName: IoWrp_Sensor
*  Content:  Io wrapper sensor module source file.
*  Category:
******************************************************************************************************************
*  Revision Management
*  yyyy.mm.dd    name              version      description
*  ----------    --------          -------      -----------------------------------
*  2025.12.31    clipping            v0001        Frist edit
******************************************************************************************************************
******************************************************************************************************************/
/* Include Headerfiles  */
/* ===================                                                  */
#include "Std_Types.h"

typedef Std_ReturnType (*WriterFunction_b)(boolean value);
typedef Std_ReturnType (*WriteValue_u8)(uint8 value);
typedef Std_ReturnType (*WriteValue_u16)(uint16 value);
typedef Std_ReturnType (*WriterFunction_32)(uint32 value);
typedef Std_ReturnType (*WriteValue_Prt)(void *value);

typedef Std_ReturnType (*ReadValue)(void *value);

typedef enum
{
  BOOLEAN = 0,
  U8,
  U16,
  U32,
  PRT
} Sensor_Type;

typedef struct
{
  Sensor_Type SensorType;
  ReadValue ReadAdcValue;
  uint8 RangeLenth;
  struct
  {
    uint16 AdcValue;
    uint8 Range;
  } AdcRange *AdcRanges;
  union
  {
    WriterFunction_b WriteBooleanValue;
    WriteValue_u8 Write8BitValue;
    WriteValue_u16 Write16BitValue;
    WriterFunction_32 Write32BitValue;
    WriteValue_Prt WritePointerValue;
  } SensorUnion;
} SensorAdc;

SensorAdc SensorAdcInstance[] = {
  { .SensorType = U8, .ReadAdcValue = NULL, .SensorUnion.Write8BitValue = WriterAary[0] },
  { .SensorType = U16, .ReadAdcValue = NULL, .SensorUnion.Write8BitValue = WriterAary[1] },
  { .SensorType = PRT, .ReadAdcValue = NULL, .SensorUnion.WritePointerValue = WriterAary[2] },
  { .SensorType = U8, .ReadAdcValue = NULL, .SensorUnion.Write8BitValue = WriterAary[0] }
};

Std_ReturnType Sensor_AdcWriteValue(const SensorAdc *SensorAdcPrt, void *Value)
{
  Std_ReturnType Ret = E_NOT_OK;
  switch (SensorAdcPrt->SensorType)
  {
    case BOOLEAN:
      Ret = SensorAdcPrt->SensorUnion.WriteBooleanValue(*(boolean *) Value);
      break;
    case U8:
      Ret = SensorAdcPrt->SensorUnion.Write8BitValue(*(uint8 *) Value);
      break;
    case U16:
      Ret = SensorAdcPrt->SensorUnion.Write16BitValue(*(uint16 *) Value);
      break;
    case U32:
      Ret = SensorAdcPrt->SensorUnion.Write32BitValue(*(uint32 *) Value);
      break;
    case PRT:
      Ret = SensorAdcPrt->SensorUnion.WritePointerValue(Value);
      break;
    default:
      Ret = E_NOT_OK;
      break;
  }

  return Ret;
}

void Sensor_AdcTransfor(const SensorAdc *SensorAdcPrt)
{
  uint16 SensorAdc;
  Std_ReturnType Ret = SensorAdcPrt->ReadAdcValue((uint16 *) &SensorAdc);
  // Initialization code for sensors can be added here
  uint8 condition = SensorAdcPrt->RangeLenth;
  uint8 index = 0;
  while (index < condition)
  {
    if (SensorAdc < SensorAdcPrt->AdcRanges[index].AdcValue)
    {
      // Some operation based on the range
      Ret |= Sensor_AdcWriteValue(SensorAdcPrt, &SensorAdcPrt->AdcRanges[index].Range);
      break;
    }
    index++;
  }
}

typedef struct
{
  Sensor_Type SensorType;
  ReadValue ReadDiValue;
  WriterFunction_b WriteBooleanValue;
} SensorDi;

SensorDi SensorDiInstance[] = {
  { .SensorType = BOOLEAN, .ReadDiValue = NULL, .WriteBooleanValue = NULL },
  { .SensorType = BOOLEAN, .ReadDiValue = NULL, .WriteBooleanValue = NULL }
};

void Sensor_DiTransfor(const SensorDi *SensorDiPrt)
{
  boolean SensorDi;
  Std_ReturnType Ret = SensorDiPrt->ReadDiValue(&SensorDi);

  Ret |= SensorDiPrt->WriteBooleanValue(SensorDi);
  break;
}

void Sensor_Mainfunction(void)
{
  // Initialization code for sensors can be added here
  SensorAdc *SensorAdcPrt = SensorAdcInstance;
  for (uint8 SensorId = 0; SensorId < sizeof(SensorAdcInstance) / sizeof(SensorAdc); SensorId++)
  {
    Sensor_AdcTransfor(&SensorAdcInstance[SensorId]);
  }

  SensorDi *SensorDiPrt = SensorDiInstance;
  for (uint8 SensorId = 0; SensorId < sizeof(SensorDiInstance) / sizeof(SensorDi); SensorId++)
  {
    Sensor_DiTransfor(&SensorDiInstance[SensorId]);
  }
}