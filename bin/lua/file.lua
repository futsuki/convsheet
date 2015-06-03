
local moduleName = "file"

evalPath[moduleName] = function (path)
   if path:match("^file:") then
      return fileget(path:sub(6))
   else
      return fileget(path)
   end
end

