// Lightweight helper: expose extractTextFromVBA(buffer) to attempt reading vbaProject.bin as latin1 text.
// This is intentionally minimal — it does not fully parse OLE streams. It provides a best-effort string of the binary content
// so the main scanner can run regex searches against it.
(function(global){
  function bufferToLatin1(buf){
    try{
      return new TextDecoder('latin1').decode(buf);
    }catch(e){
      var arr=new Uint8Array(buf);
      var s='';
      for(var i=0;i<arr.length;i++) s+=String.fromCharCode(arr[i]);
      return s;
    }
  }

  async function extractTextFromVBA(arrayBuffer){
    // best-effort: convert raw bytes to latin1 string
    return bufferToLatin1(arrayBuffer);
  }

  global.VBAExtractorHelper = { extractTextFromVBA };
})(window);