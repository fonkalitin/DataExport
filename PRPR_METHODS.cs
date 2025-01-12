using Multicad.DatabaseServices;
using Multicad.Objects;
using Multicad;
using System.Drawing;
using System.Windows;

namespace PRPR_METHODS
{
    public class PRPR_METHODS
    {
        public static void getWpfControl (object controlObject, string control_name)
        {
            //controlObject.GetType
            
        }

        public static void PRPR_HighlightByIdsList (List <McObjectId> ParObjsIdsToHighlight)
        {
            foreach (McObjectId currController in ParObjsIdsToHighlight)
            {
                McParametricObject currObjToHighlight = (McParametricObject)currController.GetObject();
                currObjToHighlight.DbEntity.Highlight(true, Color.Green);
            }
        }
        public static string PRPR_PosDescriptionAssemble(McParametricObject currParObj)
        {
            string pos_name = currParObj.Public.GetOrCreate("pos_name").AsString;
            string media_name = currParObj.Public.GetOrCreate("media_name").AsString;
            string media_wtemp = currParObj.Public.GetOrCreate("media_wtemp").AsString;
            string media_wpress = currParObj.Public.GetOrCreate("media_wpress").AsString;
            string media_tempUnits = currParObj.Public.GetOrCreate("media_temp_units").AsString;
            string media_dtemp = currParObj.Public.GetOrCreate("media_dtemp").AsString;
            string media_dpress = currParObj.Public.GetOrCreate("media_dpress").AsString;
            string media_pressUnits = currParObj.Public.GetOrCreate("media_press_units").AsString;
            string pipe_category1 = currParObj.Public.GetOrCreate("pipe_category1").AsString;
            string pipe_category2 = currParObj.Public.GetOrCreate("pipe_category2").AsString;
            string user_script = currParObj.Public.GetOrCreate("user_script").AsString;

            string result_description = "Наименование линии/оборудования: " + pos_name + ";" + "\n";
            result_description = result_description + "Наименование среды: " + media_name + ";" + "\n";
            result_description = result_description + "Рабочие параметры: " + "Tраб = " + media_wtemp + " " + media_tempUnits + "; Pраб = " + media_wpress + " " + media_pressUnits + ";" + "\n";
            result_description = result_description + "Расчетные параметры: " + "Tрасч = " + media_dtemp + " " + media_tempUnits + "; Pрасч = " + media_dpress + " " + media_pressUnits + ";" + "\n";
            result_description = result_description + "Группа среды и категория трубопровода по ГОСТ 32569-2013: " + pipe_category1 + ";" + "\n";
            result_description = result_description + "Группа среды и категория трубопровода по ТР ТС 032/2013: " + pipe_category1 + ";" + "\n";
            return result_description;
        }

        public static string PRPR_ValveDescriptionAssemble(McParametricObject currParObj)
        {
            string valveName = currParObj.Public.GetOrCreate("valve_name").AsString;
            string valveType = currParObj.Public.GetOrCreate("FitingName_Table").AsString;
            string user_script = currParObj.Public.GetOrCreate("func_user_script").AsString;
            string valve_diam = currParObj.Public.GetOrCreate("DN").AsString;
            string cond_position = currParObj.Public.GetOrCreate("normal_pos").AsString;

            string result_description = "Диаметр арматуры: " + valve_diam + ";" + "\n";
            result_description = result_description + "Позиция при условии (в зависимости от типа арматуры): " + cond_position + ";" + "\n";
            return result_description;
        }

        public static string PRPR_PrimeKipDescriptionAssemble(McParametricObject currParObj)
        {
            string kip_prefix = currParObj.Public.GetOrCreate("pos_prefix").AsString;
            string kip_func = currParObj.Public.GetOrCreate("func_kip").AsString;
            string kip_tag = currParObj.Public.GetOrCreate("kip_tag").AsString;
            string kip_suffix = currParObj.Public.GetOrCreate("pos_suffix").AsString;

            string kip_type = currParObj.Public.GetOrCreate("kip_type").AsString;
            string kip_asutype = currParObj.Public.GetOrCreate("kip_asutype").AsString;
            string kipPos_Size = currParObj.Public.GetOrCreate("pipe_curdiam").AsString;
            string kip_diam = currParObj.Public.GetOrCreate("kip_curdiam").AsString;
            string kip_signalType = currParObj.Public.GetOrCreate("kip_signal_type").AsString;
            string kip_interface = currParObj.Public.GetOrCreate("kip_interface").AsString;
            string kip_scale = currParObj.Public.GetOrCreate("kip_scale").AsString;
            string kip_exi = currParObj.Public.GetOrCreate("kip_exi").AsString;

            string result_description = "Тип КИПиА: " + kip_type + ";" + "\n";
            result_description = result_description + "Система АСУТП: " + kip_asutype + ";" + "\n";
            result_description = result_description + "Диаметр несущего трубопровода/аппарата: " + kipPos_Size + ";" + "\n";
            result_description = result_description + "Диаметр штуцера КИПиА: " + kip_diam + ";" + "\n";
            result_description = result_description + "Вид сигнала (I/O): " + kip_signalType + ";" + "\n";
            result_description = result_description + "Тип сигнала (I/O): " + kip_interface + ";" + "\n";
            result_description = result_description + "Шкала КИПиА: " + kip_scale + ";" + "\n";
            result_description = result_description + "Маркировка взрывозащиты: " + kip_exi + ";" + "\n";
            return result_description;
        }
        public static string PRPR_FuncKipDescriptionAssemble(McParametricObject currParObj)
        {
            string kip_units = currParObj.Public.GetOrCreate("kip_units").AsString;
            string kip_asutype = currParObj.Public.GetOrCreate("kip_asutype").AsString;
            double kip_H1_flag = currParObj.Public.GetOrCreate("kip_H1_flag").AsNumber;
            string kip_H1 = currParObj.Public.GetOrCreate("kip_H1").AsString;
            double kip_H2_flag = currParObj.Public.GetOrCreate("kip_H2_flag").AsNumber;
            string kip_H2 = currParObj.Public.GetOrCreate("kip_H2").AsString;
            double kip_L1_flag = currParObj.Public.GetOrCreate("kip_L1_flag").AsNumber;
            string kip_L1 = currParObj.Public.GetOrCreate("kip_L1").AsString;
            double kip_L2_flag = currParObj.Public.GetOrCreate("kip_L2_flag").AsNumber;
            string kip_L2 = currParObj.Public.GetOrCreate("kip_L2").AsString;
            double kip_HH1_flag = currParObj.Public.GetOrCreate("kip_HH1_flag").AsNumber;
            string kip_HH1 = currParObj.Public.GetOrCreate("kip_HH1").AsString;
            double kip_HH2_flag = currParObj.Public.GetOrCreate("kip_HH2_flag").AsNumber;
            string kip_HH2 = currParObj.Public.GetOrCreate("kip_HH2").AsString;
            double kip_LL1_flag = currParObj.Public.GetOrCreate("kip_LL1_flag").AsNumber;
            string kip_LL1 = currParObj.Public.GetOrCreate("kip_LL1").AsString;
            double kip_LL2_flag = currParObj.Public.GetOrCreate("kip_LL2_flag").AsNumber;
            string kip_LL2 = currParObj.Public.GetOrCreate("kip_LL2").AsString;
            string user_script = currParObj.Public.GetOrCreate("func_user_script").AsString;

            string result_description = "Система АСУТП: " + kip_asutype + ";" + "\n";

            if (kip_H1_flag == 1 || kip_H2_flag == 1 || kip_L1_flag == 1 || kip_L2_flag == 1)
            {
                result_description = result_description + "Уставки сигнализаций:" + "\n";
                if (kip_H1_flag == 1) { result_description = result_description + "H1 = " + kip_H1 + " " + kip_units + "\n"; }
                if (kip_H2_flag == 1) { result_description = result_description + "H2 = " + kip_H2 + " " + kip_units + "\n"; }
                if (kip_L1_flag == 1) { result_description = result_description + "L1 = " + kip_L1 + " " + kip_units + "\n"; }
                if (kip_L2_flag == 1) { result_description = result_description + "L2 = " + kip_L2 + " " + kip_units + "\n"; }
            }

            if (kip_HH1_flag == 1 || kip_HH2_flag == 1 || kip_LL1_flag == 1 || kip_LL2_flag == 1)
            {
                result_description = result_description + "Уставки блокировок:" + "\n";
                if (kip_HH1_flag == 1) { result_description = result_description + "HH1 = " + kip_HH1 + " " + kip_units + "\n"; }
                if (kip_HH2_flag == 1) { result_description = result_description + "HH2 = " + kip_HH2 + " " + kip_units + "\n"; }
                if (kip_LL1_flag == 1) { result_description = result_description + "LL1 = " + kip_LL1 + " " + kip_units + "\n"; }
                if (kip_LL2_flag == 1) { result_description = result_description + "LL2 = " + kip_LL2 + " " + kip_units + "\n"; }
            }

            if (user_script != "") { result_description = result_description + "Примечания проектировщика: " + "\n" + user_script + ";"; }
            return result_description;
        }

        public static string PRPR_BlockDescriptionAssemble(McParametricObject currParObj, string InfToReturn)
        {
            string curr_device = "";
            string curr_units = "";
            string curr_location = "";
            string curr_kip = "";
            string currKip_controlPoint = "";
            string curr_location_type = "";
            string pipe_name = "";

            string result_controllerDescription = "";
            List<string> blockNums = new List<string>();

            List<string> controllerDescriptions = new List<string>();
            List<string> controllerDeviceGroups = new List<string>();
            List<string> controllerDevice_blockReq = new List<string>();
            List<McObjectId> currBlockControllerHighlight = new List<McObjectId>();
            List<McObjectId> currBlockControlledHighlight = new List<McObjectId>();


            string controlledDescription = "";
            string curr_trigger1 = "";
            string curr_trigger2 = "";
            string controlled_lag = "";
            string lag_units = "";
            string subreq = "";
            string user_text = "";
            string controllerDevice_group = "";
            string block_req = "";

            ExValue currAnnotationBlockNum = currParObj.Public.GetOrCreate("block_num");
            blockNums.Add(currAnnotationBlockNum.AsString);

            ObjectFilter filter2 = ObjectFilter.Create(true);
            filter2.AddType(McParametricObject.TypeID);
            filter2.AllObjects = true;

            List<McObjectId> currBlocksIds = filter2.GetObjects();
            foreach (McObjectId currBlockId in currBlocksIds)
            {
                McParametricObject currBlockParEntity = (McParametricObject)currBlockId.GetObject();
                ExValue currBlockParEntityNum = currBlockParEntity.Public.GetOrCreate("block_num");


                if (currBlockParEntity.CommonName == "Блокировка v.0.6" && currBlockParEntityNum.AsString == currAnnotationBlockNum.AsString)
                {
                    string currBlockFunc1 = currBlockParEntity.Public.GetOrCreate("func_1").AsString;
                    curr_device = currBlockParEntity.Public.GetOrCreate("func_1").AsString + currBlockParEntity.Public.GetOrCreate("func_2").AsString;
                    curr_kip = currBlockParEntity.Public.GetOrCreate("kip_type").AsString + " " + currBlockParEntity.Public.GetOrCreate("pos").AsString;
                    curr_location_type = currBlockParEntity.Public.GetOrCreate("kip_location_type").AsString;
                    currKip_controlPoint = currBlockParEntity.Public.GetOrCreate("kip_controlPoint").AsString;
                    pipe_name = currBlockParEntity.Public.GetOrCreate("media_dzrate").AsString; // Временное решение. После актуализации поменять на нормальную переменную внутри объекта;
                    curr_location = currBlockParEntity.Public.GetOrCreate("kip_location_type").AsString + " " + currBlockParEntity.Public.GetOrCreate("kip_location").AsString;
                    curr_units = currBlockParEntity.Public.GetOrCreate("kip_units").AsString;
                    user_text = currBlockParEntity.Public.GetOrCreate("block_actions").AsString;
                    controlled_lag = currBlockParEntity.Public.GetOrCreate("controlled_lag").AsString;
                    lag_units = currBlockParEntity.Public.GetOrCreate("lag_units").AsString;
                    subreq = currBlockParEntity.Public.GetOrCreate("block_subreq").AsString;
                    curr_trigger1 = currBlockParEntity.Public.GetOrCreate("block_trigger1").AsString;
                    curr_trigger2 = currBlockParEntity.Public.GetOrCreate("block_trigger2").AsString;
                    controllerDevice_group = currBlockParEntity.Public.GetOrCreate("controller_deviceGroup").AsString;
                    block_req = currBlockParEntity.Public.GetOrCreate("block_req").AsString;

                    bool controller_flag = false;
                    bool controlled_flag = false;



                    if (currBlockFunc1 != "HS") { controller_flag = true; controlled_flag = false; } else { controller_flag = false; controlled_flag = true; }

                    if (controller_flag == true)
                    {
                        currBlockControllerHighlight.Add(currBlockId);

                        int currIndex = 0;
                        if (controllerDeviceGroups.Contains(controllerDevice_group))
                        {
                            currIndex = controllerDeviceGroups.IndexOf(controllerDevice_group);
                        }
                        else
                        {
                            controllerDeviceGroups.Add(controllerDevice_group);
                            controllerDevice_blockReq.Add(block_req);
                            currIndex = controllerDeviceGroups.IndexOf(controllerDevice_group);
                        }

                        if (controllerDescriptions.Count > currIndex)
                        {
                            controllerDescriptions[currIndex] = controllerDescriptions[currIndex] + "\n" + "     При значении поз." + curr_device;
                        }
                        else
                        {
                            controllerDescriptions.Add("     При значении поз." + curr_device);
                        }

                        if (curr_trigger1 == "HH1") { controllerDescriptions[currIndex] = controllerDescriptions[currIndex] + "\n" + "     HH1>=" + currBlockParEntity.Public.GetOrCreate("kip_HH1").AsString + " " + curr_units; }
                        if (curr_trigger1 == "HH2") { controllerDescriptions[currIndex] = controllerDescriptions[currIndex] + "\n" + "     HH2>=" + currBlockParEntity.Public.GetOrCreate("kip_HH2").AsString + " " + curr_units; }
                        if (curr_trigger2 == "LL1") { controllerDescriptions[currIndex] = controllerDescriptions[currIndex] + "\n" + "     LL1<=" + currBlockParEntity.Public.GetOrCreate("kip_LL1").AsString + " " + curr_units; }
                        if (curr_trigger2 == "LL2") { controllerDescriptions[currIndex] = controllerDescriptions[currIndex] + "\n" + "     LL2<=" + currBlockParEntity.Public.GetOrCreate("kip_LL2").AsString + " " + curr_units; }

                        controllerDescriptions[currIndex] = controllerDescriptions[currIndex] + "\n" + "     Место измерения: " + curr_location;

                        if (currKip_controlPoint != "") { controllerDescriptions[currIndex] = controllerDescriptions[currIndex] + "\n" + "     Точка измерения: " + currKip_controlPoint; }
                        if (curr_location_type == "Линия") { controllerDescriptions[currIndex] = controllerDescriptions[currIndex] + "\n" + "     Название линии: " + pipe_name + ";"; }
                    }

                    if (controlled_flag == true)
                    {
                        currBlockControlledHighlight.Add(currBlockId);

                        controlledDescription = controlledDescription + "\n" + curr_kip + ";" + "\n" + "     Сигналы: ";

                        if (curr_trigger1 == "H1") { controlledDescription = controlledDescription + currBlockParEntity.Public.GetOrCreate("kip_HH1").AsString + ";"; }
                        if (curr_trigger1 == "H2") { controlledDescription = controlledDescription + currBlockParEntity.Public.GetOrCreate("kip_HH2").AsString + ";"; }
                        if (curr_trigger2 == "L1") { controlledDescription = controlledDescription + currBlockParEntity.Public.GetOrCreate("kip_LL1").AsString + ";"; }
                        if (curr_trigger2 == "L2") { controlledDescription = controlledDescription + currBlockParEntity.Public.GetOrCreate("kip_LL2").AsString + ";"; }
                        if (controlled_lag != "Отсутствует") { controlledDescription = controlledDescription + "\n" + "     Задержка выполнения:" + controlled_lag + " " + lag_units + ";"; }
                        if (subreq != "(пусто)") { controlledDescription = controlledDescription + "\n" + "     [только при условии: " + subreq + "];"; }

                        controlledDescription = controlledDescription + "\n" + "     Место размещения исполнительного механизма:" + curr_location + ";";

                        if (curr_location_type == "Линия") { controlledDescription = controlledDescription + "\n" + "     Название линии: " + pipe_name + ";"; }
                        if (user_text != "" && user_text != "(пусто)") { controlledDescription = controlledDescription + "\n" + "ПОЛЬЗОВАТЕЛЬСКИЙ ТЕКСТ " + user_text + ";"; }
                    }
                }
            }

            foreach (string currGroup_description in controllerDescriptions)
            {
                result_controllerDescription = result_controllerDescription + "\n" + "[Группа датчиков " + (controllerDescriptions.IndexOf(currGroup_description) + 1) + ". Схема срабатывания: " + controllerDevice_blockReq[controllerDescriptions.IndexOf(currGroup_description)] + "]" + "\n" + currGroup_description;
            }

            

            if (InfToReturn == "controllerDescription")
            {
                return result_controllerDescription;
            }
            else if (InfToReturn == "controlledDescription")
            {
                return controlledDescription;
            }
            else
            {
                return null;
            }
        }
        public static void PRPR_BlockDescriptionAssemble(McParametricObject currParObj, out string controllersDescription, out string controlsDescription, out List <McObjectId> controllersIds, out List<McObjectId> controlledIds)
        {

            string curr_device = "";
            string curr_units = "";
            string curr_location = "";
            string curr_kip = "";
            string currKip_controlPoint = "";
            string curr_location_type = "";
            string pipe_name = "";

            string result_controllerDescription = "";
            List<string> blockNums = new List<string>();

            List<string> controllerDescriptions = new List<string>();
            List<string> controllerDeviceGroups = new List<string>();
            List<string> controllerDevice_blockReq = new List<string>();
            List<McObjectId> currBlockControllerHighlight = new List<McObjectId>();
            List<McObjectId> currBlockControlledHighlight = new List<McObjectId>();


            string controlledDescription = "";
            string curr_trigger1 = "";
            string curr_trigger2 = "";
            string controlled_lag = "";
            string lag_units = "";
            string subreq = "";
            string user_text = "";
            string controllerDevice_group = "";
            string block_req = "";

            ExValue currAnnotationBlockNum = currParObj.Public.GetOrCreate("block_num");
            blockNums.Add(currAnnotationBlockNum.AsString);

            ObjectFilter filter2 = ObjectFilter.Create(true);
            filter2.AddType(McParametricObject.TypeID);
            filter2.AllObjects = true;

            List<McObjectId> currBlocksIds = filter2.GetObjects();
            foreach (McObjectId currBlockId in currBlocksIds)
            {
                McParametricObject currBlockParEntity = (McParametricObject)currBlockId.GetObject();
                ExValue currBlockParEntityNum = currBlockParEntity.Public.GetOrCreate("block_num");


                if (currBlockParEntity.CommonName == "Блокировка v.0.6" && currBlockParEntityNum.AsString == currAnnotationBlockNum.AsString)
                {
                    string currBlockFunc1 = currBlockParEntity.Public.GetOrCreate("func_1").AsString;
                    curr_device = currBlockParEntity.Public.GetOrCreate("func_1").AsString + currBlockParEntity.Public.GetOrCreate("func_2").AsString;
                    curr_kip = currBlockParEntity.Public.GetOrCreate("kip_type").AsString + " " + currBlockParEntity.Public.GetOrCreate("pos").AsString;
                    curr_location_type = currBlockParEntity.Public.GetOrCreate("kip_location_type").AsString;
                    currKip_controlPoint = currBlockParEntity.Public.GetOrCreate("kip_controlPoint").AsString;
                    pipe_name = currBlockParEntity.Public.GetOrCreate("media_dzrate").AsString; // Временное решение. После актуализации поменять на нормальную переменную внутри объекта;
                    curr_location = currBlockParEntity.Public.GetOrCreate("kip_location_type").AsString + " " + currBlockParEntity.Public.GetOrCreate("kip_location").AsString;
                    curr_units = currBlockParEntity.Public.GetOrCreate("kip_units").AsString;
                    user_text = currBlockParEntity.Public.GetOrCreate("block_actions").AsString;
                    controlled_lag = currBlockParEntity.Public.GetOrCreate("controlled_lag").AsString;
                    lag_units = currBlockParEntity.Public.GetOrCreate("lag_units").AsString;
                    subreq = currBlockParEntity.Public.GetOrCreate("block_subreq").AsString;
                    curr_trigger1 = currBlockParEntity.Public.GetOrCreate("block_trigger1").AsString;
                    curr_trigger2 = currBlockParEntity.Public.GetOrCreate("block_trigger2").AsString;
                    controllerDevice_group = currBlockParEntity.Public.GetOrCreate("controller_deviceGroup").AsString;
                    block_req = currBlockParEntity.Public.GetOrCreate("block_req").AsString;

                    bool controller_flag = false;
                    bool controlled_flag = false;



                    if (currBlockFunc1 != "HS") { controller_flag = true; controlled_flag = false; } else { controller_flag = false; controlled_flag = true; }

                    if (controller_flag == true)
                    {
                        currBlockControllerHighlight.Add(currBlockId);

                        int currIndex = 0;
                        if (controllerDeviceGroups.Contains(controllerDevice_group))
                        {
                            currIndex = controllerDeviceGroups.IndexOf(controllerDevice_group);
                        }
                        else
                        {
                            controllerDeviceGroups.Add(controllerDevice_group);
                            controllerDevice_blockReq.Add(block_req);
                            currIndex = controllerDeviceGroups.IndexOf(controllerDevice_group);
                        }

                        if (controllerDescriptions.Count > currIndex)
                        {
                            controllerDescriptions[currIndex] = controllerDescriptions[currIndex] + "\n" + "     При значении поз." + curr_device;
                        }
                        else
                        {
                            controllerDescriptions.Add("     При значении поз." + curr_device);
                        }

                        if (curr_trigger1 == "HH1") { controllerDescriptions[currIndex] = controllerDescriptions[currIndex] + "\n" + "     HH1>=" + currBlockParEntity.Public.GetOrCreate("kip_HH1").AsString + " " + curr_units; }
                        if (curr_trigger1 == "HH2") { controllerDescriptions[currIndex] = controllerDescriptions[currIndex] + "\n" + "     HH2>=" + currBlockParEntity.Public.GetOrCreate("kip_HH2").AsString + " " + curr_units; }
                        if (curr_trigger2 == "LL1") { controllerDescriptions[currIndex] = controllerDescriptions[currIndex] + "\n" + "     LL1<=" + currBlockParEntity.Public.GetOrCreate("kip_LL1").AsString + " " + curr_units; }
                        if (curr_trigger2 == "LL2") { controllerDescriptions[currIndex] = controllerDescriptions[currIndex] + "\n" + "     LL2<=" + currBlockParEntity.Public.GetOrCreate("kip_LL2").AsString + " " + curr_units; }

                        controllerDescriptions[currIndex] = controllerDescriptions[currIndex] + "\n" + "     Место измерения: " + curr_location;

                        if (currKip_controlPoint != "") { controllerDescriptions[currIndex] = controllerDescriptions[currIndex] + "\n" + "     Точка измерения: " + currKip_controlPoint; }
                        if (curr_location_type == "Линия") { controllerDescriptions[currIndex] = controllerDescriptions[currIndex] + "\n" + "     Название линии: " + pipe_name + ";"; }
                    }

                    if (controlled_flag == true)
                    {
                        currBlockControlledHighlight.Add(currBlockId);

                        controlledDescription = controlledDescription + "\n" + curr_kip + ";" + "\n" + "     Сигналы: ";

                        if (curr_trigger1 == "H1") { controlledDescription = controlledDescription + currBlockParEntity.Public.GetOrCreate("kip_HH1").AsString + ";"; }
                        if (curr_trigger1 == "H2") { controlledDescription = controlledDescription + currBlockParEntity.Public.GetOrCreate("kip_HH2").AsString + ";"; }
                        if (curr_trigger2 == "L1") { controlledDescription = controlledDescription + currBlockParEntity.Public.GetOrCreate("kip_LL1").AsString + ";"; }
                        if (curr_trigger2 == "L2") { controlledDescription = controlledDescription + currBlockParEntity.Public.GetOrCreate("kip_LL2").AsString + ";"; }
                        if (controlled_lag != "Отсутствует") { controlledDescription = controlledDescription + "\n" + "     Задержка выполнения:" + controlled_lag + " " + lag_units + ";"; }
                        if (subreq != "(пусто)") { controlledDescription = controlledDescription + "\n" + "     [только при условии: " + subreq + "];"; }

                        controlledDescription = controlledDescription + "\n" + "     Место размещения исполнительного механизма:" + curr_location + ";";

                        if (curr_location_type == "Линия") { controlledDescription = controlledDescription + "\n" + "     Название линии: " + pipe_name + ";"; }
                        if (user_text != "" && user_text != "(пусто)") { controlledDescription = controlledDescription + "\n" + "ПОЛЬЗОВАТЕЛЬСКИЙ ТЕКСТ " + user_text + ";"; }
                    }
                }
            }

            foreach (string currGroup_description in controllerDescriptions)
            {
                result_controllerDescription = result_controllerDescription + "\n" + "[Группа датчиков " + (controllerDescriptions.IndexOf(currGroup_description) + 1) + ". Схема срабатывания: " + controllerDevice_blockReq[controllerDescriptions.IndexOf(currGroup_description)] + "]" + "\n" + currGroup_description;
            }

            controllersIds = currBlockControllerHighlight;
            controlledIds = currBlockControlledHighlight;

            controllersDescription = result_controllerDescription;
            controlsDescription = controlledDescription;
        }
        
        
        public static List<List<ExValue>> CollectParObjData(string parObjCommonName, List<string> attributesList, int collectArea)
        {
            var attributesObjsExValues = new List<List<ExValue>>();

            ObjectFilter filter = ObjectFilter.Create(false);
            filter.AddType(McParametricObject.TypeID);
            filter.AllObjects = true;
            if (collectArea == 0) { filter.SetCurrentSheet(); } else if (collectArea == 1) { filter.SetCurrentDocument(); }
            List<McObjectId> filterIds = filter.GetObjects();

            foreach (McObjectId id in filterIds)
            {
                McParametricObject currParObj = (McParametricObject)id.GetObject();
                if (currParObj.CommonName == parObjCommonName)
                {
                    var attributesObjExValues = new List<ExValue>();
                    foreach (string attributeName in attributesList)
                    {
                        attributesObjExValues.Add(currParObj.Public.GetOrCreate(attributeName));
                    }
                    attributesObjsExValues.Add(attributesObjExValues);
                }
            }
            return attributesObjsExValues;
        }
    }
}
