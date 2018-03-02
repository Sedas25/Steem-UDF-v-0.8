#include <Steem Websocket.au3>
#include <Array.au3>
_Steem_Start_Node("api.steemit.com")
;<==========================================
; Read The Content From the Post
$sPost = _Steem_Get_Content("utopian-io", "utopian-is-expanding-collaboration-with-browserstack-crowdin-disasterhack-and-more")
;<==========================================
;<==========================================
; Display The Content in a 2D Array
_ArrayDisplay($sPost,"Content Display")
;<==========================================

;<==========================================
; Read The Last 500 Votes From utopian-io
$sVotes = _Steem_Get_Account_Votes("utopian-io", 500)
;<==========================================
;<==========================================
; Display The Votes in a 2D Array
_ArrayDisplay($sVotes,"Votes Display")
;<==========================================

_Steem_Close_Node()