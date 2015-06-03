
local moduleName = "protobuf"
local push = table.insert

checkSheet[moduleName] = function ()
   local l = sheet[1]
   return
      (l[1] or ""):lower() == "message" and
      (l[2] or ""):lower() == "assign" and
      (l[3] or ""):lower() == "type" and
      (l[4] or ""):lower() == "name" and
      (l[5] or ""):lower() == "value" and
      (l[6] or ""):lower() == "attribute"
end

local function clamp(s)
   local r = (s or ""):gsub("^[　 \t\n]+", ""):gsub("[　 \t\n]+$", "")
   return r
end

local function emptyOr(s, val)
   if s == "" then
      return val
   else
      return s
   end
end

applySheet[moduleName] = function ()
   local t = {}
   local w = function(s) push(t, s) end
   local indent = ""
   local ws = function(l)
       push(t, indent)
       push(t, clamp(table.concat(l, " ")))
       push(t, "\n")
   end
   
   local defaultLine = nil
   local currentMessage = nil
   for i=2, #sheet do
      local l = sheet[i]
      if l[1]:lower() == "default" then
         defaultLine = l
      else
         if l[1]:lower() == "#declare" then
            if l[5] ~= "" then
               ws{l[2], l[3], l[4], "=", l[5], ";"}
            else
               ws{l[2], l[3], l[4], ";"}
            end
         elseif l[4] == "" then
         
         else
            if l[1] ~= "" and currentMessage ~= l[1] then
               indent = ""
               if currentMessage then
                  ws{"}"}
                  ws{}
               end
               currentMessage = l[1]
               ws{"message", l[1], "{"}
               indent = "  "
            end
            w(indent)
            w(emptyOr(l[2], defaultLine[2]))
            w" "
            w(emptyOr(l[3], defaultLine[3]))
            w" "
            w(emptyOr(l[4], defaultLine[4]))
            w(" = ")
            w(l[5])
            w" "
            w(l[6])
            w(";\n")
         end
      end
   end
   if currentMessage then
      w("}\n")
   end

   --[[
      {filefull}
      {file} <- without extension
      {ext}
      {dir}
      {sheet}
   ]]
   return {
      ["{file}_{sheet}.proto"] = table.concat(t)
   }
end



