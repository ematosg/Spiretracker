/*
 * Spire Online Client (Supabase)
 *
 * Optional helper module. It does not alter existing local behavior until
 * app.js imports and uses it.
 */

(function (global) {
  const state = {
    supabase: null,
    ready: false
  };

  function normalizeRpcScalar(data, fieldNames = []) {
    if (data == null) return data;
    if (typeof data === 'string' || typeof data === 'number' || typeof data === 'boolean') return data;
    if (Array.isArray(data)) {
      if (!data.length) return null;
      return normalizeRpcScalar(data[0], fieldNames);
    }
    if (typeof data === 'object') {
      for (const key of fieldNames) {
        if (data[key] != null) return data[key];
      }
      const values = Object.values(data);
      if (values.length === 1) return values[0];
    }
    return data;
  }

  function init(config) {
    const url = config && config.url;
    const key = config && config.anonKey;
    if (!url || !key) {
      throw new Error('Missing Supabase config (url/anonKey).');
    }
    if (!global.supabase || !global.supabase.createClient) {
      throw new Error('Supabase client library not loaded.');
    }
    state.supabase = global.supabase.createClient(url, key);
    state.ready = true;
    return state.supabase;
  }

  function client() {
    if (!state.ready || !state.supabase) {
      throw new Error('Online client not initialized.');
    }
    return state.supabase;
  }

  async function signUp(email, password, username, accountType) {
    const sb = client();
    const { data, error } = await sb.auth.signUp({
      email,
      password,
      options: {
        data: {
          username,
          account_type: accountType
        }
      }
    });
    if (error) throw error;
    return data;
  }

  async function signIn(email, password) {
    const sb = client();
    const { data, error } = await sb.auth.signInWithPassword({ email, password });
    if (error) throw error;
    return data;
  }

  async function signOut() {
    const sb = client();
    const { error } = await sb.auth.signOut();
    if (error) throw error;
  }

  async function currentUser() {
    const sb = client();
    const { data, error } = await sb.auth.getUser();
    if (error) throw error;
    return data.user;
  }

  async function createCampaign(name) {
    const sb = client();
    const { data, error } = await sb.rpc('create_campaign', { p_name: name || 'New Campaign' });
    if (error) throw error;
    return normalizeRpcScalar(data, ['create_campaign']);
  }

  async function generateInviteCode(campaignId, roleToGrant = 'player', maxUses = 1, expiresMinutes = 1440) {
    const sb = client();
    const { data, error } = await sb.rpc('generate_invite_code', {
      p_campaign_id: campaignId,
      p_role_to_grant: roleToGrant,
      p_max_uses: maxUses,
      p_expires_minutes: expiresMinutes
    });
    if (error) throw error;
    return normalizeRpcScalar(data, ['generate_invite_code', 'code']);
  }

  async function joinCampaignWithCode(code) {
    const sb = client();
    const { data, error } = await sb.rpc('join_campaign_with_code', {
      p_code: code
    });
    if (error) throw error;
    return normalizeRpcScalar(data, ['join_campaign_with_code', 'campaign_id']);
  }

  async function revokeInviteCode(campaignId, code) {
    const sb = client();
    const normalized = String(code || '').trim().toUpperCase();
    if (!normalized) throw new Error('Invite code is required.');
    const { data, error } = await sb
      .from('invite_codes')
      .update({ revoked: true })
      .eq('campaign_id', campaignId)
      .eq('code', normalized)
      .select('id,code,revoked')
      .maybeSingle();
    if (error) throw error;
    return data;
  }

  async function listMyCampaigns() {
    const sb = client();
    const { data, error } = await sb
      .from('campaign_members')
      .select('role,campaigns!inner(id,name,owner_user_id,data,created_at,updated_at)')
      .order('joined_at', { ascending: true });
    if (error) throw error;
    return (data || []).map(row => ({
      role: row.role,
      campaign: row.campaigns
    }));
  }

  async function saveCampaignData(campaignId, campaignData) {
    const sb = client();
    const { data, error } = await sb
      .from('campaigns')
      .update({ data: campaignData })
      .eq('id', campaignId)
      .select('id,updated_at')
      .single();
    if (error) throw error;
    return data;
  }

  global.SpireOnlineClient = {
    init,
    signUp,
    signIn,
    signOut,
    currentUser,
    createCampaign,
    generateInviteCode,
    joinCampaignWithCode,
    revokeInviteCode,
    listMyCampaigns,
    saveCampaignData
  };
})(typeof window !== 'undefined' ? window : globalThis);
