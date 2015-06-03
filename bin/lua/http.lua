
local moduleName = "http"

evalPath[moduleName] = function (path)
   if path:match("^http://") then
      return httpget(path)
   else
      return nil
   end
end

